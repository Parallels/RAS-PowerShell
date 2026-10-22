#!/usr/bin/env python3
"""Mock HPE Morpheus VME appliance.

Payload shapes follow real VME responses:

  - list entries and GET ?details=false are the same lean object: connectionInfo[]
    carries the IP, top-level interfaces[].macAddress is null, no MAC anywhere;
  - details=true adds containerDetails[].server.interfaces[].macAddress;
  - MAC and OS live on /api/servers objects (which also carry credential-ish fields
    that must never be logged);
  - clones carry config.provisionOpts.clonedInstanceId (lineage);
  - delete goes running -> removing -> 404;
  - clone/start/stop/delete return {"success":true} with no processIds.

Extras for tests:
  - HTTP/1.1 keep-alive and a connection counter (proves connection reuse);
  - per-endpoint request counters at /_ctl/stats (proves call budgets);
  - fault injection: status codes, dropped connections, 429 Retry-After, a clone that
    is accepted but whose response is delayed past the client timeout, a failed
    listing page, an empty listing, a forced 401;
  - a fixture of N filler instances (POST /_ctl/reset?count=N).

Run:  python mock_vme.py [port]        # https://127.0.0.1:8444 by default
"""
import ssl, json, re, sys, os, time, uuid, threading, http.server
from urllib.parse import urlparse, parse_qs

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8444
TOKEN = "mock-token-abc"
REFRESH = "mock-refresh-xyz"
CREDS = {"admin": "password"}
HERE = os.path.dirname(os.path.abspath(__file__))
LOGFILE = os.path.join(HERE, f"mock-{PORT}.log")

DEFAULT_KNOBS = {
    "clone_delay": 2.0,        # s before a cloned instance appears
    "provision_delay": 2.0,    # s after appearing before status=running
    "mac_delay": 1.0,          # s after appearing before its server has a MAC
    "ip_delay": 3.0,           # s after running before it has an IP
    "delete_delay": 1.0,       # s in 'removing' before 404
    "action_delay": 1.0,       # s for start -> running
    "phrase_enabled": True,    # False: phrase= is ignored and matches nothing
    "activity_enabled": True,
    "clone_honours_mac": True, # networkInterfaces[0].macAddress applied to the clone
    "clone_honours_network": True, # networkInterfaces[0].network.id applied to the clone
    "clone_hang_once": 0.0,    # accept the next clone, then delay the response this long
    "pad_bytes": 400,          # filler per instance, to make listings heavier
    "token_expires_in": 86400,
    "clone_final_status": "running",  # 'stopped' = provisioned powered off (provisionPoweredOff)
    # Stuck-clone model: the clone is accepted but no instance ever
    # appears; its 'provision' process sits unfinished on the SOURCE's process history
    # until /_ctl/ghost ends it.
    "clone_never_appears": False,
    "ghost_display_name": True,  # False: the ghost process's displayName is not the clone name
    # Instance listing (not phrase): max=-1 returns everything.
    "list_max_cap": 0,         # >0: no response carries more rows than this (a capping VME build)
    "reject_max_all": False,   # True: max<=0 is answered 400 (a VME that refuses max=-1)
    "list_delay": 0.0,         # s added to every listing request (slow appliance)
    # VME 9.1 template versions: snapshot -> process; linked-clone
    # is synchronous and makes a 'Synced' virtual image named like the snapshot, whose
    # locations[].externalDiskId is the snapshot's externalId; deleting the snapshot
    # deletes the image.
    "snapshot_delay": 2.0,     # s before a snapshot (and its process) is complete
    "snapshot_fails": False,   # the snapshot process ends 'failed'
    "linked_clone_fails": False,  # PUT linked-clone answers 400 {"success":false} (9.0.x stub)
    "image_bucket": "Synced",  # filterType bucket the linked-clone image lands in
    "image_delay": 0.0,        # s after the linked-clone call before the image is listed
    "image_links_snapshot": True,  # False: externalDiskId does not match the snapshot's externalId
    "build_version": "9.1.0",
    # Right after a stop VME answers 'stopping', and a listing tens of seconds later can
    # still say 'running'. >0: stop -> 'stopping', then 'running' again after 0.3 s,
    # then 'stopped' after this many seconds.
    "stop_lag": 0.0,
}

lock = threading.Lock()
_logf = open(LOGFILE, "a", buffering=1)


def log(msg):
    _logf.write(time.strftime("%H:%M:%S ") + "[mock-vme] " + msg + "\n")


def iso(ts):
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(ts))


def mac_for(sid):
    return "52:54:00:%02X:%02X:%02X" % ((sid >> 16) & 255, (sid >> 8) & 255, sid & 255)


def new_instance(iid, name, status="running", ip=None, cloned_from=None, mac=None, os_name="windows 11 64-bit",
                  net_id=5, net_name="Compute VLAN 100", srv_net_id=None, srv_net_name=None):
    sid = iid + 4                                        # server id != instance id, like the samples
    # net_id/net_name: the INSTANCE's own interfaces[].network (GET /api/instances) --
    # Morpheus does NOT keep this in sync when a VM's
    # network is changed underneath it.
    # srv_net_id/srv_net_name: the SERVER's interfaces[].network (GET /api/servers/{id})
    # -- the ground truth. Independently patchable via /_ctl/instance to model the drift.
    if srv_net_id is None:
        srv_net_id, srv_net_name = net_id, net_name
    return {
        "id": iid, "uuid": str(uuid.uuid4()), "name": name, "status": status,
        "ip": ip, "server_id": sid, "server_uuid": str(uuid.uuid4()),
        "mac": mac or mac_for(sid), "mac_visible": True, "os": os_name,
        "net_id": net_id, "net_name": net_name,
        "srv_net_id": srv_net_id, "srv_net_name": srv_net_name,
        # the network VME's clone backup/snapshot actually preserves when it ignores the
        # requested override (clone_honours_network=false) -- the ground truth (server-
        # level) at creation time, not the possibly-stale instance-level display field
        "backup_net_id": srv_net_id, "backup_net_name": srv_net_name,
        "cloned_from": cloned_from, "image_id": None, "busy": False, "percent": None,
        "created": time.time(), "updated": time.time(), "removing_since": None,
    }


def seed(count=41):
    insts = {}
    insts[19] = new_instance(19, "tmpl-w11", "stopped", None)
    insts[20] = new_instance(20, "plain-vm", "running", "192.168.1.51")
    insts[21] = new_instance(21, "stopped-vm", "stopped", "192.168.1.52")   # stale IP on a stopped VM
    insts[22] = new_instance(22, "failed-vm", "failed", None)
    insts[23] = new_instance(23, "linklocal-vm", "running", "169.254.10.10")
    insts[24] = new_instance(24, "busy-vm", "running", "0.0.0.0")
    for n in range(100, 100 + count):
        insts[n] = new_instance(n, f"filler-{n}", "running", f"10.1.{(n // 250) % 250}.{n % 250 + 1}", cloned_from=19)
    return insts


state = {}


def reset(count=41):
    state.clear()
    state.update({
        "instances": seed(count), "next_id": 10000, "knobs": dict(DEFAULT_KNOBS),
        "fail": {},            # key -> [remaining, code]; code 0 = drop connection
        "empty_list_next": 0,
        "fail_list_page_offset": None,
        "force_401": 0,
        "stats": {}, "connections": 0,
        "last_clone_body": None,
        "tokens": {TOKEN},
        "cancelled_processes": set(),
        "ghosts": [],          # clones that never appeared: {id, src, name, status, created}
        "snapshots": {},       # id -> {id, iid, name, externalId, created, done_at, failed}
        "images": {},          # id -> virtual image (Library)
        "next_snap": 300, "next_image": 900,
        "last_create_body": None, "last_revert": None, "last_image_update": None,
        "last_power": None,    # {iid, action, query}: the query string of the last start/stop/restart/suspend
    })


reset()


def later(delay, fn):
    def run():
        time.sleep(delay)
        with lock:
            fn()
    threading.Thread(target=run, daemon=True).start()


def lean(i):
    """GET /api/instances list entry == GET /api/instances/{id}?details=false."""
    k = state["knobs"]
    out = {
        "id": i["id"], "uuid": i["uuid"], "accountId": 1,
        "instanceType": {"id": 10, "code": "mvm", "category": "cloud", "name": "HVM"},
        "group": {"id": 1, "name": "Demo Cloud Group"},
        "cloud": {"id": 1, "name": "HVM-Cloud", "type": "standard"},
        "servers": [i["server_id"]],
        "connectionInfo": [{"ip": i["ip"], "port": None, "name": None}] if i["ip"] else [],
        "layout": {"id": 32, "name": "Single HVM", "provisionTypeCode": "kvm"},
        "plan": {"id": 23, "code": "kvm-vm-8192", "name": "2 CPU, 8GB Memory"},
        "name": i["name"], "displayName": i["name"],
        "config": dict({"imageId": i["image_id"] or 217, "resourcePoolId": "pool-1", "poolProviderType": "mvm", "uefi": True,
                        "secureBoot": True, "tpm": True, "attachVirtIODrivers": True, "nestedVirtualization": False,
                        "noAgent": False, "kvmHostId": None},
                       provisionOpts={"clonedInstanceId": i["cloned_from"], "imageId": 217} if i["cloned_from"] else {}),
        "maxCores": 2, "coresPerSocket": 1, "maxMemory": 8589934592,
        "volumes": [
            {"id": 199, "name": "root", "rootVolume": True, "size": 60, "volumeCategory": "disk", "datastoreId": "auto"},
            {"id": 200, "name": "CD Drive", "rootVolume": False, "size": 0, "volumeCategory": "cd", "datastoreId": None},
        ],
        "interfaces": [{"id": 5, "network": {"id": i["net_id"], "name": i["net_name"]}, "ipAddress": None, "macAddress": None}],
        "evars": [{"name": "MVM_IP", "value": "0.0.0.0"}],
        "tags": [], "labels": [],
        "dateCreated": iso(i["created"]), "lastUpdated": iso(i["updated"]),
        "status": i["status"], "statusMessage": None, "statusPercent": i["percent"], "statusEta": None,
        "isBusy": i["busy"],
        "stats": {"filler": "x" * k["pad_bytes"]},
    }
    return out


def proc_status(i):
    """An instance's own provision process: unfinished only while VME is building it
    (a stopped template's own process is long complete)."""
    if i["status"] == "failed":
        return "failed"
    if i["status"] in ("provisioning", "pending", "deploying") or i["busy"]:
        return "running"
    return "complete"


def server_obj(i):
    return {
        "id": i["server_id"], "uuid": i["server_uuid"], "name": i["name"],
        "powerState": "on" if i["status"] == "running" else "off",
        "internalIp": i["ip"], "externalIp": i["ip"], "osType": "windows",
        "serverOs": {"id": 23, "code": "windows.11.64", "name": i["os"]},
        "interfaces": [{"id": 151, "name": "eth0", "primaryInterface": True, "ipAddress": i["ip"],
                        "macAddress": i["mac"] if i["mac_visible"] else None,
                        "network": {"id": i["srv_net_id"], "name": i["srv_net_name"]}}],
        "instance": {"id": i["id"], "name": i["name"]},
        "sshUsername": "admin", "sshPassword": "ssh-secret-value", "apiKey": "server-api-key-value",
        "computeServerType": {"code": "morpheusWindowsVm"},
    }


def detailed(i):
    out = lean(i)
    out["containerDetails"] = [{"id": i["id"] + 1, "server": server_obj(i)}]
    return out


def snap_status(s):
    if time.time() < s["done_at"]:
        return "creating"
    return "failed" if s["failed"] else "complete"


def snapshot_obj(s):
    """Shape of GET /api/snapshots/{id} on VME 9.1."""
    return {"id": s["id"], "name": s["name"], "description": s.get("description"), "externalId": s["externalId"],
            "status": snap_status(s), "snapshotType": "vm", "snapshotCreated": iso(s["created"]),
            "zone": {"id": 1, "name": "HVM-Cloud"}, "datastore": None, "parentSnapshot": None,
            "snapshotFiles": [{"id": s["id"] * 10, "name": "vda", "type": "file",
                               "externalId": f"vm-{s['iid']}/disk-{s['iid']}-snap-{s['id']}"}],
            "currentlyActive": True, "memorySnapshot": False, "forExport": False, "forBackup": False,
            "dateCreated": iso(s["created"])}


def image_visible(img):
    return time.time() >= img["visible_at"]


def image_obj(img):
    """Shape of GET /api/virtual-images/{id} on VME 9.1."""
    return {"id": img["id"], "name": img["name"], "description": None, "labels": img["labels"], "imageType": "qcow2",
            "userUploaded": False, "userDefined": True, "systemImage": False, "minDisk": None, "uefi": True,
            "tpm": True, "secureBoot": True, "linkedClone": True, "externalId": f"img-ext-{img['iid']}",
            "visibility": "private", "config": {}, "volumes": [], "networkInterfaces": [], "tags": [],
            "locations": [{"id": img["id"] + 50000, "externalId": f"img-ext-{img['iid']}", "externalDiskId": img["disk"],
                           "imageName": img["name"], "refType": "ComputeServerGroup", "refId": 2, "diskIndex": 0}],
            "dateCreated": iso(img["created"]), "lastUpdated": iso(img["created"]), "status": None}


def stat_key(method, path, qs):
    key = method + " " + re.sub(r"/\d+", "/{id}", path)
    if path == "/api/instances":
        key += " phrase" if "phrase" in qs else (" name" if "name" in qs else " list")
    if "details" in qs and qs["details"][0] == "true":
        key += " details"
    return key


class Handler(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def setup(self):
        super().setup()
        with lock:
            state["connections"] += 1

    def log_message(self, fmt, *args):
        log(fmt % args)

    def _send(self, code, obj, headers=None):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _err(self, code, msg, errors=None):
        body = {"success": False, "msg": msg}
        if errors:
            body["errors"] = errors
        return self._send(code, body)

    def _body(self):
        n = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(n) if n else b""
        if not raw:
            return {}
        try:
            return json.loads(raw)
        except Exception:
            return {k: v[0] for k, v in parse_qs(raw.decode("utf-8")).items()}

    def _count(self, method, path, qs):
        k = stat_key(method, path, qs)
        with lock:
            state["stats"][k] = state["stats"].get(k, 0) + 1

    def _fault(self, method, path):
        """Injected faults. Key 'METHOD /path/{id}' or 'METHOD /exact/path'."""
        keys = [f"{method} {path}", f"{method} " + re.sub(r"/\d+", "/{id}", path)]
        with lock:
            hit = None
            for k in keys:
                f = state["fail"].get(k)
                if f and f[0] > 0:
                    f[0] -= 1
                    hit = (k, f[1], f[2] if len(f) > 2 else None)
                    break
        if not hit:
            return False
        k, code, extra = hit
        log(f"FAULT {k} -> {code}")
        if code == 0:
            self.close_connection = True
            try:
                self.connection.shutdown(2)
            except Exception:
                pass
            return True
        if code == 429:
            return self._send(429, {"success": False, "msg": "Too many requests"}, {"Retry-After": str(extra or 1)}) or True
        self._err(code, f"Injected failure {code}")
        return True

    def _auth(self, path):
        if not path.startswith("/api/"):
            return True
        with lock:
            if state["force_401"] > 0:
                state["force_401"] -= 1
                self._err(401, "Unauthorized (forced)")
                return False
            ok = self.headers.get("Authorization", "").split(None, 1)
            if len(ok) == 2 and ok[0].upper() == "BEARER" and ok[1] in state["tokens"]:
                return True
        self._err(401, "Unauthorized")
        return False

    # ------------------------------------------------------------------ control
    def _control(self, method, path, qs, body):
        args = dict(body)
        for k, v in qs.items():
            args[k] = v[0]
        if path == "/_ctl/reset":
            with lock:
                reset(int(args.get("count", 41)))
            return self._send(200, {"success": True})
        if path == "/_ctl/stats":
            with lock:
                return self._send(200, {"requests": dict(state["stats"]), "connections": state["connections"]})
        if path == "/_ctl/stats/reset":
            with lock:
                state["stats"] = {}
                state["connections"] = 0
            return self._send(200, {"success": True})
        if path == "/_ctl/set":
            with lock:
                for k, v in args.items():
                    if k not in state["knobs"]:
                        return self._err(400, f"unknown knob {k}")
                    cur = state["knobs"][k]
                    state["knobs"][k] = (str(v).lower() in ("1", "true", "yes")) if isinstance(cur, bool) else type(cur)(v)
            return self._send(200, {"success": True, "knobs": state["knobs"]})
        if path == "/_ctl/fail":
            with lock:
                state["fail"][args["key"]] = [int(args.get("count", 1)), int(args.get("code", 500)), int(args.get("extra", 0))]
            return self._send(200, {"success": True})
        if path == "/_ctl/list":
            with lock:
                if "empty" in args:
                    state["empty_list_next"] = int(args["empty"])
                if "fail_offset" in args:
                    state["fail_list_page_offset"] = int(args["fail_offset"])
            return self._send(200, {"success": True})
        if path == "/_ctl/force401":
            with lock:
                state["force_401"] = int(args.get("count", 1))
                if args.get("revoke") == "1":
                    state["tokens"] = set()
            return self._send(200, {"success": True})
        if path == "/_ctl/instance":
            # add/patch an instance directly: id, name, status, ip, cloned_from, delete=1
            with lock:
                iid = int(args["id"])
                if args.get("delete") == "1":
                    state["instances"].pop(iid, None)
                    return self._send(200, {"success": True})
                i = state["instances"].get(iid) or new_instance(iid, args.get("name", f"vm-{iid}"))
                for k in ("name", "status"):
                    if k in args:
                        i[k] = args[k]
                if "ip" in args:
                    i["ip"] = args["ip"] or None
                if "cloned_from" in args:
                    i["cloned_from"] = int(args["cloned_from"]) if args["cloned_from"] else None
                if "busy" in args:
                    i["busy"] = args["busy"] == "1"
                if "net_id" in args:
                    i["net_id"] = int(args["net_id"])
                if "net_name" in args:
                    i["net_name"] = args["net_name"]
                if "srv_net_id" in args:
                    i["srv_net_id"] = int(args["srv_net_id"])
                if "srv_net_name" in args:
                    i["srv_net_name"] = args["srv_net_name"]
                state["instances"][iid] = i
            return self._send(200, {"success": True, "instance": {"id": iid, "uuid": i["uuid"], "mac": i["mac"]}})
        if path == "/_ctl/ghost":
            # add=1&src=&name=: a never-appeared clone (as if accepted earlier);
            # otherwise end its provision process: name, status (failed/complete)
            with lock:
                if args.get("add") == "1":
                    state["ghosts"].append({"id": 7000 + len(state["ghosts"]), "src": int(args["src"]), "name": args["name"],
                                            "status": "running", "created": time.time()})
                    return self._send(200, {"success": True})
                hit = [g for g in state["ghosts"] if g["name"] == args["name"]]
                for g in hit:
                    g["status"] = args.get("status", "failed")
            return self._send(200, {"success": True, "ended": len(hit)})
        if path == "/_ctl/state":
            with lock:
                return self._send(200, {"instances": {str(k): {"name": v["name"], "status": v["status"], "mac": v["mac"], "cloned_from": v["cloned_from"], "image_id": v["image_id"]} for k, v in state["instances"].items()},
                                        "last_clone_body": state["last_clone_body"],
                                        "last_create_body": state["last_create_body"],
                                        "last_revert": state["last_revert"],
                                        "last_image_update": state["last_image_update"],
                                        "last_power": state["last_power"],
                                        "snapshots": {str(k): {"iid": s["iid"], "name": s["name"], "status": snap_status(s)} for k, s in state["snapshots"].items()},
                                        "images": {str(k): {"name": m["name"], "disk": m["disk"], "labels": m["labels"], "bucket": m["bucket"]} for k, m in state["images"].items()},
                                        "cancelled_processes": sorted(state["cancelled_processes"])})
        if path == "/_ctl/snapshot":
            # add a snapshot directly (as if made earlier): iid, name
            with lock:
                sid = state["next_snap"]
                state["next_snap"] += 1
                state["snapshots"][sid] = {"id": sid, "iid": int(args["iid"]), "name": args["name"], "externalId": f"snap-ext-{sid}",
                                           "created": time.time(), "done_at": time.time(), "failed": False}
            return self._send(200, {"success": True, "id": sid})
        return self._err(404, f"no control path {path}")

    # ---------------------------------------------------------------------- GET
    def do_GET(self):
        u = urlparse(self.path)
        path, qs = u.path, parse_qs(u.query)
        if path.startswith("/_ctl/"):
            return self._control("GET", path, qs, {})
        self._count("GET", path, qs)
        if not self._auth(path) or self._fault("GET", path):
            return

        if path == "/api/whoami":
            return self._send(200, {"user": {"id": 1, "username": "admin"}})
        if path == "/api/user-settings":
            return self._send(200, {"user": {"id": 1}, "accessTokens": [
                {"clientId": "morph-api", "maskedAccessToken": TOKEN[:7] + "********", "expiration": iso(time.time() + 30 * 86400)}]})

        if path == "/api/instances":
            mx = int(qs.get("max", ["25"])[0])
            off = int(qs.get("offset", ["0"])[0])
            if "phrase" not in qs:
                if state["knobs"]["list_delay"] > 0:
                    time.sleep(state["knobs"]["list_delay"])
                if mx <= 0 and state["knobs"]["reject_max_all"]:
                    return self._err(400, "max must be a positive integer")
            with lock:
                if mx <= 0:
                    mx = len(state["instances"])
                if state["knobs"]["list_max_cap"] > 0 and "phrase" not in qs:
                    mx = min(mx, state["knobs"]["list_max_cap"])
                if state["fail_list_page_offset"] is not None and off == state["fail_list_page_offset"] and "phrase" not in qs:
                    state["fail_list_page_offset"] = None
                    return self._err(500, "Injected page failure")
                if state["empty_list_next"] > 0 and off == 0 and "phrase" not in qs:
                    state["empty_list_next"] -= 1
                    return self._send(200, {"instances": [], "meta": {"offset": 0, "max": mx, "size": 0, "total": 0}})
                items = [i for _, i in sorted(state["instances"].items())]
                if "phrase" in qs:
                    ph = qs["phrase"][0].lower()
                    items = [i for i in items if ph in i["name"].lower()] if state["knobs"]["phrase_enabled"] else []
                # name= is deliberately ignored (fix6: a real appliance did exactly that)
                total = len(items)
                page = [lean(i) for i in items[off:off + mx]]
            return self._send(200, {"instances": page, "meta": {"offset": off, "max": mx, "size": len(page), "total": total}})

        m = re.match(r"^/api/instances/(\d+)$", path)
        if m:
            iid = int(m.group(1))
            details = qs.get("details", ["true"])[0].lower() != "false"
            with lock:
                i = state["instances"].get(iid)
                if i is None:
                    return self._err(404, f"Instance not found for id {iid}")
                return self._send(200, {"instance": detailed(i) if details else lean(i)})

        if path == "/api/servers":
            mx = int(qs.get("max", ["25"])[0])
            off = int(qs.get("offset", ["0"])[0])
            with lock:
                servers = [server_obj(i) for _, i in sorted(state["instances"].items())]
                servers.append({"id": 999999, "name": "unmanaged-vm", "interfaces": [{"macAddress": "52:54:00:FF:FF:FF"}], "instance": None,
                                "serverOs": {"name": "linux"}, "sshPassword": "x"})
                page = servers[off:off + mx]
                total = len(servers)
            return self._send(200, {"servers": page, "meta": {"offset": off, "max": mx, "size": len(page), "total": total}})

        m = re.match(r"^/api/servers/(\d+)$", path)
        if m:
            sid = int(m.group(1))
            with lock:
                for i in state["instances"].values():
                    if i["server_id"] == sid:
                        return self._send(200, {"server": server_obj(i), "stats": {}})
            return self._err(404, f"Server not found for id {sid}")

        if path == "/api/processes":
            iid = int(qs.get("instanceId", ["0"])[0])
            with lock:
                i = state["instances"].get(iid)
                if i is None:
                    return self._send(200, {"processes": [], "meta": {"total": 0}})
                k = state["knobs"]
                age = time.time() - i["created"]
                deploy_done = age >= k["provision_delay"] / 2
                st = proc_status(i)
                proc = {"id": 5000 + iid, "processType": {"code": "provision", "name": "provision"}, "status": st,
                        "displayName": i["name"], "instanceId": iid,
                        "startDate": iso(i["created"]), "error": "disk copy failed" if i["status"] == "failed" else None,
                        "events": [{"processType": {"code": "provisionResources", "name": "Prepare Resources"}, "status": "complete"},
                                   {"processType": {"code": "provisionDeploy", "name": "Deploy Instance"}, "status": "complete" if deploy_done else "running"}]}
                procs = [proc]
                for g in state["ghosts"]:
                    if g["src"] == iid:
                        procs.append({"id": g["id"], "processType": {"code": "provision", "name": "provision"}, "status": g["status"],
                                      "displayName": g["name"] if k["ghost_display_name"] else f"container_{g['id']}",
                                      "instanceId": iid, "startDate": iso(g["created"]), "error": None, "events": []})
            return self._send(200, {"processes": procs, "meta": {"total": len(procs)}})

        if path == "/api/activity":
            name = qs.get("name", [""])[0]
            with lock:
                rows = []
                if state["knobs"]["activity_enabled"]:
                    rows = [{"_id": str(uuid.uuid4()), "success": True, "activityType": "Provisioning", "name": i["name"],
                             "message": "Instance created", "objectType": "Instance", "objectId": i["id"], "ts": iso(i["created"])}
                            for i in state["instances"].values() if i["name"] == name]
            return self._send(200, {"activity": rows, "meta": {"size": len(rows), "total": len(rows)}})

        if path == "/api/ping":
            return self._send(200, {"success": True, "buildVersion": state["knobs"]["build_version"], "applianceUrl": "https://127.0.0.1"})

        m = re.match(r"^/api/instances/(\d+)/snapshots$", path)
        if m:
            iid = int(m.group(1))
            with lock:
                if iid not in state["instances"]:
                    return self._err(404, f"Instance not found for id {iid}")
                snaps = [snapshot_obj(s) for _, s in sorted(state["snapshots"].items()) if s["iid"] == iid]
            return self._send(200, {"snapshots": snaps})

        m = re.match(r"^/api/processes/(\d+)$", path)
        if m:
            pid = int(m.group(1))
            with lock:
                s = state["snapshots"].get(pid - 8000)
                if s is None:
                    return self._err(404, f"Process not found for id {pid}")
                st = snap_status(s)
                pst = {"creating": "running"}.get(st, st)
                return self._send(200, {"process": {"id": pid, "processType": {"code": "snapshot", "name": "snapshot"}, "status": pst,
                                                    "percent": 100 if pst == "complete" else 40, "instanceId": s["iid"],
                                                    "error": "qemu snapshot failed" if pst == "failed" else None, "events": []}})

        if path == "/api/virtual-images":
            ft = qs.get("filterType", ["User"])[0]
            name = qs.get("name", [None])[0]
            with lock:
                imgs = [image_obj(m) for _, m in sorted(state["images"].items())
                        if image_visible(m) and (ft == "All" or ft == m["bucket"]) and (name is None or m["name"] == name)]
            return self._send(200, {"virtualImages": imgs, "meta": {"size": len(imgs), "total": len(imgs)}})

        m = re.match(r"^/api/virtual-images/(\d+)$", path)
        if m:
            with lock:
                img = state["images"].get(int(m.group(1)))
                if img is None:
                    return self._err(404, "Virtual Image not found")
                return self._send(200, {"virtualImage": image_obj(img)})

        return self._err(404, f"no such endpoint {path}")

    # --------------------------------------------------------------------- POST
    def do_POST(self):
        u = urlparse(self.path)
        path, qs = u.path, parse_qs(u.query)
        body = self._body()
        if path.startswith("/_ctl/"):
            return self._control("POST", path, qs, body)
        self._count("POST", path, qs)
        if self._fault("POST", path):
            return
        if path == "/oauth/token":
            args = {k: v[0] for k, v in qs.items()}
            args.update(body)
            grant = args.get("grant_type")
            with lock:
                if grant == "password":
                    if CREDS.get(args.get("username")) != args.get("password"):
                        return self._err(400, "Invalid credentials")
                elif grant == "refresh_token":
                    if args.get("refresh_token") != REFRESH:
                        return self._err(400, "Invalid refresh token")
                else:
                    return self._err(400, "unsupported grant")
                state["tokens"].add(TOKEN)
                exp = state["knobs"]["token_expires_in"]
            return self._send(200, {"access_token": TOKEN, "refresh_token": REFRESH, "expires_in": int(exp), "token_type": "bearer"})

        m = re.match(r"^/api/processes/(\d+)/cancel$", path)
        if m:
            pid = int(m.group(1))
            iid = pid - 5000  # matches the synthetic process id in do_GET's /api/processes
            with lock:
                i = state["instances"].get(iid)
                if i is None:
                    return self._err(404, f"Process not found for id {pid}")
                st = proc_status(i)
                if st not in ("new", "queued", "pending", "running"):
                    return self._err(400, "Process is not cancelable")
                state["cancelled_processes"].add(pid)
                log(f"process {pid} (instance {iid}) cancelled")
            return self._send(200, {"success": True})

        if path == "/api/instances":
            # addInstance from a linked-clone virtual image (config.imageId)
            if not self._auth(path):
                return
            with lock:
                inst = body.get("instance") or {}
                cfg = body.get("config") or {}
                name = inst.get("name")
                img = state["images"].get(int(cfg.get("imageId") or 0))
                if not name:
                    return self._err(400, "Name is required", {"name": "Name is required"})
                if img is None:
                    return self._err(400, "Invalid image", {"imageId": "Image not found"})
                if not body.get("zoneId") or not inst.get("plan") or not inst.get("layout"):
                    return self._err(400, "Missing placement", {"zoneId": "required"})
                state["last_create_body"] = body
                nid = state["next_id"]
                state["next_id"] += 1
                k = dict(state["knobs"])
                nics = body.get("networkInterfaces") or []
                req_mac = nics[0].get("macAddress") if nics else None
                req_net = (nics[0].get("network") or {}).get("id") if nics else None
                i = new_instance(nid, name, "provisioning", None, mac=req_mac,
                                 net_id=req_net or 5, net_name=f"VLAN-{req_net}" if req_net else "Compute VLAN 100")
                i["image_id"] = img["id"]
                i["busy"] = True
                i["percent"] = 10
                state["instances"][nid] = i
                out = lean(i)
                log(f"addInstance {nid} ({name}) from image {img['id']}, provisioning")

            def done():
                j = state["instances"].get(nid)
                if j and j["status"] == "provisioning":
                    j["status"] = k["clone_final_status"]
                    j["busy"] = False
                    j["percent"] = None

                    def ip():
                        if nid in state["instances"]:
                            state["instances"][nid]["ip"] = f"10.8.{(nid // 250) % 250}.{nid % 250 + 1}"
                    later(k["ip_delay"], ip)
            later(k["provision_delay"], done)
            return self._send(200, {"instance": out})

        return self._err(404, f"no such endpoint {path}")

    # ---------------------------------------------------------------------- PUT
    def do_PUT(self):
        u = urlparse(self.path)
        path, qs = u.path, parse_qs(u.query)
        body = self._body()
        if path.startswith("/_ctl/"):
            return self._control("PUT", path, qs, body)
        self._count("PUT", path, qs)
        if not self._auth(path) or self._fault("PUT", path):
            return

        m = re.match(r"^/api/instances/(\d+)/(start|stop|restart|suspend)$", path)
        if m:
            iid, action = int(m.group(1)), m.group(2)
            with lock:
                state["last_power"] = {"iid": iid, "action": action, "query": u.query}
                i = state["instances"].get(iid)
                if i is None:
                    return self._err(404, f"Instance not found for id {iid}")
                if i["busy"]:
                    return self._err(400, "Instance is busy")
                if action in ("start", "restart"):
                    i["status"] = "starting"

                    def fin(i=i):
                        if i["status"] == "starting":
                            i["status"] = "running"
                    later(state["knobs"]["action_delay"], fin)
                elif action == "stop" and state["knobs"]["stop_lag"] > 0:
                    i["status"] = "stopping"

                    def back(i=i):
                        if i["status"] == "stopping":
                            i["status"] = "running"

                    def off(i=i):
                        if i["status"] in ("stopping", "running"):
                            i["status"] = "stopped"
                    later(0.3, back)
                    later(state["knobs"]["stop_lag"], off)
                elif action == "stop":
                    i["status"] = "stopped"
                else:
                    i["status"] = "suspended"
                i["updated"] = time.time()
            return self._send(200, {"success": True})

        m = re.match(r"^/api/instances/(\d+)/clone$", path)
        if m:
            src = int(m.group(1))
            with lock:
                s = state["instances"].get(src)
                if s is None:
                    return self._err(404, f"Instance not found for id {src}")
                name = body.get("name")
                nics = body.get("networkInterfaces") or []
                if not name:
                    return self._err(400, "Name is required", {"name": "Name is required"})
                if not nics:
                    return self._err(400, "Invalid network", {"networkInterfaces": "Invalid network"})
                if not body.get("volumes"):
                    return self._err(400, "Invalid volume", {"volumes": "Invalid volume"})
                state["last_clone_body"] = body
                nid = state["next_id"]
                state["next_id"] += 1
                k = dict(state["knobs"])
                hang = k["clone_hang_once"]
                state["knobs"]["clone_hang_once"] = 0.0
                req_mac = nics[0].get("macAddress")
                req_net = (nics[0].get("network") or {}).get("id")
                src_backup_net_id, src_backup_net_name = s["backup_net_id"], s["backup_net_name"]

            def appear():
                if req_net and k["clone_honours_network"]:
                    net_id, net_name = req_net, f"VLAN-{req_net}"
                else:
                    net_id, net_name = src_backup_net_id, src_backup_net_name  # ignored: clone kept the backup's network
                i = new_instance(nid, name, "provisioning", None, cloned_from=src,
                                 mac=req_mac if (req_mac and k["clone_honours_mac"]) else None,
                                 net_id=net_id, net_name=net_name)
                i["busy"] = True
                i["percent"] = 10
                i["mac_visible"] = False
                state["instances"][nid] = i
                log(f"clone {nid} ({name}) visible, provisioning")

                def mac_now():
                    if nid in state["instances"]:
                        state["instances"][nid]["mac_visible"] = True
                later(k["mac_delay"], mac_now)

                def done():
                    j = state["instances"].get(nid)
                    if j and j["status"] == "provisioning":
                        j["status"] = k["clone_final_status"]
                        j["busy"] = False
                        j["percent"] = None

                        def ip():
                            if nid in state["instances"]:
                                state["instances"][nid]["ip"] = f"10.9.{(nid // 250) % 250}.{nid % 250 + 1}"
                        later(k["ip_delay"], ip)
                later(k["provision_delay"], done)
            if k["clone_never_appears"]:
                with lock:
                    state["ghosts"].append({"id": 7000 + len(state["ghosts"]), "src": src, "name": name,
                                            "status": "running", "created": time.time()})
                log(f"clone of {src} as {name} accepted -- stuck, no instance will appear")
            else:
                later(k["clone_delay"], appear)
            if hang > 0:
                time.sleep(hang)
            return self._send(200, {"success": True})

        m = re.match(r"^/api/instances/(\d+)/snapshot$", path)
        if m:
            iid = int(m.group(1))
            with lock:
                if iid not in state["instances"]:
                    return self._err(404, f"Instance not found for id {iid}")
                snap = body.get("snapshot") or {}
                sid = state["next_snap"]
                state["next_snap"] += 1
                k = state["knobs"]
                state["snapshots"][sid] = {"id": sid, "iid": iid, "name": snap.get("name") or f"vm-{iid}.{int(time.time())}",
                                           "description": snap.get("description"), "externalId": f"snap-ext-{sid}",
                                           "created": time.time(), "done_at": time.time() + k["snapshot_delay"], "failed": k["snapshot_fails"]}
            return self._send(200, {"success": True, "processIds": [8000 + sid]})

        m = re.match(r"^/api/instances/(\d+)/linked-clone/(\d+)$", path)
        if m:
            iid, sid = int(m.group(1)), int(m.group(2))
            with lock:
                k = state["knobs"]
                if k["linked_clone_fails"]:
                    return self._send(400, {"success": False})
                s = state["snapshots"].get(sid)
                if s is None or s["iid"] != iid:
                    return self._err(404, f"Snapshot not found for id {sid}")
                mid = state["next_image"]
                state["next_image"] += 1
                state["images"][mid] = {"id": mid, "iid": iid, "name": s["name"], "labels": [], "bucket": k["image_bucket"],
                                        "disk": s["externalId"] if k["image_links_snapshot"] else f"disk-{mid}",
                                        "snapshot": sid, "created": time.time(), "visible_at": time.time() + k["image_delay"]}
            return self._send(200, {"success": True})

        m = re.match(r"^/api/instances/(\d+)/revert-snapshot/(\d+)$", path)
        if m:
            iid, sid = int(m.group(1)), int(m.group(2))
            with lock:
                s = state["snapshots"].get(sid)
                if s is None or s["iid"] != iid:
                    return self._err(404, f"Snapshot not found for id {sid}")
                state["last_revert"] = {"iid": iid, "sid": sid, "name": s["name"]}
            return self._send(200, {"success": True})

        m = re.match(r"^/api/virtual-images/(\d+)$", path)
        if m:
            with lock:
                img = state["images"].get(int(m.group(1)))
                if img is None:
                    return self._err(404, "Virtual Image not found")
                vi = body.get("virtualImage") or {}
                if "labels" in vi:
                    img["labels"] = list(vi["labels"])
                if "name" in vi:
                    img["name"] = vi["name"]
                state["last_image_update"] = body
                return self._send(200, {"success": True, "virtualImage": image_obj(img)})

        return self._err(404, f"no such endpoint {path}")

    # ------------------------------------------------------------------- DELETE
    def do_DELETE(self):
        u = urlparse(self.path)
        path, qs = u.path, parse_qs(u.query)
        if path.startswith("/_ctl/"):
            return self._control("DELETE", path, qs, {})
        self._count("DELETE", path, qs)
        if not self._auth(path) or self._fault("DELETE", path):
            return
        m = re.match(r"^/api/instances/(\d+)$", path)
        if m:
            iid = int(m.group(1))
            with lock:
                i = state["instances"].get(iid)
                if i is None:
                    return self._err(404, f"Instance not found for id {iid}")
                i["status"] = "removing"
                d = state["knobs"]["delete_delay"]

            def gone():
                state["instances"].pop(iid, None)
                log(f"instance {iid} removed")
            later(d, gone)
            return self._send(200, {"success": True})
        m = re.match(r"^/api/snapshots/(\d+)$", path)
        if m:
            sid = int(m.group(1))
            with lock:
                s = state["snapshots"].get(sid)
                if s is None:
                    return self._err(404, f"Snapshot not found for id {sid}")
                d = state["knobs"]["delete_delay"]

            def gone():
                # Deleting the snapshot deletes its linked-clone image.
                state["snapshots"].pop(sid, None)
                for mid in [k2 for k2, img in state["images"].items() if img["snapshot"] == sid]:
                    state["images"].pop(mid, None)
                log(f"snapshot {sid} removed with its image(s)")
            later(d, gone)
            return self._send(200, {"success": True})
        m = re.match(r"^/api/virtual-images/(\d+)$", path)
        if m:
            with lock:
                if state["images"].pop(int(m.group(1)), None) is None:
                    return self._err(404, "Virtual Image not found")
            return self._send(200, {"success": True})
        return self._err(404, f"no such endpoint {path}")


def main():
    server = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    server.daemon_threads = True
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=os.path.join(HERE, "cert.pem"), keyfile=os.path.join(HERE, "key.pem"))
    server.socket = ctx.wrap_socket(server.socket, server_side=True)
    log(f"listening on https://127.0.0.1:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
