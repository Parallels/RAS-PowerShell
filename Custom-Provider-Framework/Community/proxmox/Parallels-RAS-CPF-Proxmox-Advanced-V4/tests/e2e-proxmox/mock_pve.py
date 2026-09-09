#!/usr/bin/env python3
import ssl, json, re, sys, os, time, threading, http.server
from urllib.parse import urlparse, parse_qs

PORT = 8443

state = {
    "next_id": 300,
    "vms": {
        "153": {"name": "template", "node": "n1", "status": "stopped", "template": 1, "tags": ""},
        "100": {"name": "always-on", "node": "n1", "status": "running", "template": 0, "agent": True,
                "mac": "BC:24:11:00:00:64", "ip": "10.0.0.100", "tags": ""},
        "101": {"name": "agentless", "node": "n1", "status": "running", "template": 0, "agent": False, "tags": ""},
        "102": {"name": "admin-excluded", "node": "n1", "status": "running", "template": 0, "agent": False,
                "tags": "rasExclude"},
    },
    "tasks": {},  # upid -> {"kind":"clone"/"start"/"stop"/"destroy", "vmid":..., "done_at": ts, "result": "OK"}
    "clone_delay": 1.0,   # seconds before a clone task reports done
}
lock = threading.Lock()


def new_upid(kind, vmid):
    n = len(state["tasks"]) + 1
    return f"UPID:n1:{n:08X}:00000000:00000000:{kind}:{vmid}:test@pve!auto:"


class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("[mock-pve] " + (fmt % args) + "\n")

    def _send(self, code, obj, reason=None):
        body = json.dumps(obj).encode("utf-8")
        if reason is not None:
            self.send_response(code, reason)
        else:
            self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b""
        if not raw:
            return {}
        try:
            return json.loads(raw)
        except Exception:
            return {k: v[0] for k, v in parse_qs(raw.decode("utf-8")).items()}

    def do_GET(self):
        path = urlparse(self.path).path
        qs = parse_qs(urlparse(self.path).query)

        if path == "/api2/json/version":
            return self._send(200, {"data": {"version": "8.2.4"}})

        if path == "/api2/json/cluster/nextid":
            with lock:
                vmid = str(state["next_id"])
            return self._send(200, {"data": vmid})

        if path == "/api2/json/cluster/resources":
            with lock:
                data = []
                for vmid, vm in state["vms"].items():
                    if vm.get("hidden"):
                        continue
                    entry = {"type": "qemu", "vmid": int(vmid), "name": vm["name"],
                             "node": vm["node"], "status": vm["status"], "template": vm["template"]}
                    if vm.get("tags"):
                        entry["tags"] = vm["tags"]
                    data.append(entry)
            return self._send(200, {"data": data})

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/config$", path)
        if m:
            vmid = m.group(2)
            with lock:
                vm = state["vms"].get(vmid)
            if vm is None:
                return self._send(500, {"errors": {"vmid": "vm not found"}})
            # 'template' is real Proxmox config state. 'clone_full' is not a real
            # Proxmox field -- it is this mock's own record of the 'full' value a
            # clone call actually used (after applying the real "non-template source
            # is always full" rule below), exposed here purely so the E2E suite can
            # assert linked-vs-full without inferring it from timing.
            return self._send(200, {"data": {"tags": vm.get("tags", ""), "template": vm.get("template", 0),
                                              "clone_full": vm.get("clone_full")}})

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/status/current$", path)
        if m:
            vmid = m.group(2)
            with lock:
                vm = state["vms"].get(vmid)
            if vm is None:
                return self._send(500, {"errors": {"vmid": "vm not found"}})
            return self._send(200, {"data": {"status": vm["status"], "qmpstatus": vm["status"]}})

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/agent/network-get-interfaces$", path)
        if m:
            vmid = m.group(2)
            with lock:
                vm = state["vms"].get(vmid, {})
            if not vm.get("agent"):
                return self._send(500, {"errors": {"agent": "QEMU guest agent is not running"}})
            iface = {"hardware-address": vm.get("mac", "AA:BB:CC:00:00:00"),
                     "ip-addresses": [{"ip-address-type": "ipv4", "ip-address": vm.get("ip", "10.0.0.1")}]}
            return self._send(200, {"data": {"result": [iface]}})

        m = re.match(r"^/api2/json/nodes/([^/]+)/tasks/([^/]+)/status$", path)
        if m:
            from urllib.parse import unquote
            upid = unquote(m.group(2))
            with lock:
                task = state["tasks"].get(upid)
            if task is None:
                return self._send(200, {"data": {"status": "stopped", "exitstatus": "OK"}})
            done = time.time() >= task["done_at"]
            if done:
                return self._send(200, {"data": {"status": "stopped", "exitstatus": task["result"]}})
            return self._send(200, {"data": {"status": "running"}})

        return self._send(404, {"errors": {"path": "not found"}})

    def do_POST(self):
        path = urlparse(self.path).path
        body = self._read_body()

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/clone$", path)
        if m:
            node, srcid = m.group(1), m.group(2)
            newid = str(body.get("newid"))
            name = body.get("name", f"VM {newid}")
            requested_full = body.get("full")
            with lock:
                src = state["vms"].get(srcid, {})
                # Real Proxmox full-clones copy the source VM's config -- tags
                # included -- to the new VM. Mirrored here so the E2E run
                # exercises the same tag-inheritance-then-strip path a real
                # clone does.
                src_tags = src.get("tags", "")
                # The one rule that matters for linked clones, straight from the PVE
                # schema: cloning a normal (non-template) VM is ALWAYS a full copy
                # regardless of the 'full' flag -- no error, no warning. A provider
                # that skips checking the source's template flag before asking for
                # full=0 would have Proxmox silently do a full copy anyway; this must
                # reproduce that silence so the provider-side guard is the thing
                # actually tested, not assumed.
                src_is_template = src.get("template") == 1
                if not src_is_template:
                    effective_full = 1
                elif requested_full is None:
                    effective_full = 0
                else:
                    effective_full = int(requested_full)
                # 'storage'/'format' are full-clone-only and rejected by Proxmox on a
                # linked clone.
                if effective_full == 0 and ("storage" in body or "format" in body):
                    return self._send(500, {"errors": {"storage": "parameter 'storage' not allowed for linked clones"}},
                                       reason="parameter 'storage' not allowed for linked clones")
                state["vms"][newid] = {"name": f"VM {newid}", "node": node, "status": "stopped",
                                        "template": 0, "agent": True, "final_name": name,
                                        "hidden": True, "tags": src_tags, "clone_full": effective_full}
                upid = new_upid("qmclone", newid)
                state["tasks"][upid] = {"kind": "clone", "vmid": newid, "done_at": time.time() + state["clone_delay"], "result": "OK"}
            # reveal the VM in cluster/resources (with placeholder name) shortly
            # after the clone call returns, and rename it once the task finishes --
            # mirrors real Proxmox behaviour closely enough for this smoke test.
            def reveal():
                time.sleep(0.05)
                with lock:
                    state["vms"][newid]["hidden"] = False
                time.sleep(state["clone_delay"])
                with lock:
                    state["vms"][newid]["name"] = name
            threading.Thread(target=reveal, daemon=True).start()
            return self._send(200, {"data": upid})

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/template$", path)
        if m:
            vmid = m.group(2)
            with lock:
                vm = state["vms"].get(vmid)
                if vm is None:
                    return self._send(500, {"errors": {"vmid": "not found"}})
                if vm.get("template") == 1:
                    # Real Proxmox's own rejection -- see MAINTENANCE-MODE.md and
                    # Handle-GuestConvert's race-catch branch for this exact string.
                    return self._send(500, {"errors": {"template": "you can't convert a template to a template"}},
                                       reason="you can't convert a template to a template")
                vm["template"] = 1
                upid = new_upid("qmtemplate", vmid)
                state["tasks"][upid] = {"kind": "template", "vmid": vmid, "done_at": time.time(), "result": "OK"}
            return self._send(200, {"data": upid})

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/status/(start|stop|shutdown)$", path)
        if m:
            vmid, action = m.group(2), m.group(3)
            with lock:
                vm = state["vms"].get(vmid)
                if vm is None:
                    return self._send(500, {"errors": {"vmid": "not found"}})
                # simulate PVE's real clone-lock rejection
                for t in state["tasks"].values():
                    if t["kind"] == "clone" and t["vmid"] == vmid and time.time() < t["done_at"]:
                        return self._send(595, {"errors": {"lock": "can't lock file - got timeout"}},
                                           reason="can't lock file - got timeout")
                new_status = "running" if action == "start" else "stopped"
                vm["status"] = new_status
                upid = new_upid(f"qm{action}", vmid)
                state["tasks"][upid] = {"kind": action, "vmid": vmid, "done_at": time.time(), "result": "OK"}
            return self._send(200, {"data": upid})

        return self._send(404, {"errors": {"path": "not found"}})

    def do_DELETE(self):
        path = urlparse(self.path).path
        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)$", path)
        if m:
            vmid = m.group(2)
            with lock:
                state["vms"].pop(vmid, None)
                upid = new_upid("qmdestroy", vmid)
                state["tasks"][upid] = {"kind": "destroy", "vmid": vmid, "done_at": time.time(), "result": "OK"}
            return self._send(200, {"data": upid})
        return self._send(404, {"errors": {"path": "not found"}})

    def do_PUT(self):
        path = urlparse(self.path).path
        body = self._read_body()

        m = re.match(r"^/api2/json/nodes/([^/]+)/qemu/(\d+)/config$", path)
        if m:
            vmid = m.group(2)
            with lock:
                vm = state["vms"].get(vmid)
                if vm is None:
                    return self._send(500, {"errors": {"vmid": "not found"}})
                if "tags" in body:
                    vm["tags"] = body["tags"]
                if "template" in body:
                    vm["template"] = int(body["template"])
            return self._send(200, {"data": None})

        return self._send(404, {"errors": {"path": "not found"}})


def main():
    server = http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    # Resolved against this file, not the CWD, so the mock can be launched from anywhere.
    here = os.path.dirname(os.path.abspath(__file__))
    ctx.load_cert_chain(certfile=os.path.join(here, "cert.pem"), keyfile=os.path.join(here, "key.pem"))
    server.socket = ctx.wrap_socket(server.socket, server_side=True)
    print(f"[mock-pve] listening on https://127.0.0.1:{PORT}", file=sys.stderr)
    server.serve_forever()


if __name__ == "__main__":
    main()
