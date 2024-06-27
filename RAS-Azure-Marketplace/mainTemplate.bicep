param localAdminUser string = 'rasdmin'
@secure()
param localAdminPassword string
param vnetName string
param vnetNewOrExisting string
param vnetkAddressPrefix string
param vnetResourceGroup string
param subnetName string
param subnetAddressPrefix string
param prefixCBName string
param prefixSGName string
param msName string
param numberofCBs int = 2
param numberofSGs int = 2
param vmSkuCB string = 'Standard_D4s_v5'
param vmSkuSG string = 'Standard_D4s_v5'
param vmSkuMS string = 'Standard_D2s_v5'
param vmOSCB string = '2022-Datacenter'
param vmOSSG string = '2022-Datacenter'
param vmOSMS string = '2022-Datacenter'
param lbName string = 'ras-lb-01'
param RasAdminsGroupAD string
param location string = resourceGroup().location
param keyVaultName string = 'pras-kv-01'
param azureADTenantID string = tenant().tenantId
param adminAccountObjectID string
param domainJoinUserName string
@secure()
param domainJoinPassword string
param domainName string
param ouPath string
param customerUsageAttributionID string
param providerSelection string
param providerName string = 'Provider'
param providerAppRegistrationName string = 'ras-app'

param rasVersion string = 'evergreen'
param customURLRAS string = 'evergreen'

var appPublisherName = 'Parallels'
var appProductName = 'parallelsrasprod'
var enabledForDeployment = true
var enabledForTemplateDeployment = true
var enabledForDiskEncryption = true
var enableRbacAuthorization = false
var softDeleteRetentionInDays = 90
var accessPolicies = [
  {
    tenantId: azureADTenantID
    objectId: adminAccountObjectID
    permissions: {
      keys: ['Get', 'List']
      secrets: ['Get', 'List', 'Set']
      certificates: ['Get', 'List']
    }
  }
]
var downloadURLRAS = rasVersion == 'customURL' ? customURLRAS : 'https://download.parallels.com/ras/latest/RASInstaller.msi'
var assetLocation = 'https://raw.githubusercontent.com/Parallels/RAS-PowerShell/master/RAS-Azure-Marketplace/'
var configurationScriptRAS = 'RAS_Azure_MP_Install.ps1'
var registerScriptRAS = 'RAS_Azure_MP_Register.ps1'
var prereqScript = 'RAS_Azure_MP_prereq.ps1'
var connectionBrokerPrimaryScript = 'RAS_Azure_MP_Primary_CB.ps1'
var privateIPAllocationMethod = 'Dynamic'
var vnetId = {
  new: virtualNetwork.id
  existing: resourceId(vnetResourceGroup, 'Microsoft.Network/virtualNetworks', vnetName)
}
var subnetId = '${vnetId[vnetNewOrExisting]}/subnets/${subnetName}'
var domainJoinOptions = 3
var lbSkuName = 'Standard'
var licenseType = 1

resource virtualNetwork 'Microsoft.Network/virtualNetworks@2023-06-01' =
  if (vnetNewOrExisting == 'new') {
    name: vnetName
    location: location
    properties: {
      addressSpace: {
        addressPrefixes: [vnetkAddressPrefix]
      }
      subnets: [
        {
          name: subnetName
          properties: {
            addressPrefix: subnetAddressPrefix
          }
        }
      ]
    }
  }

resource ConnectionBrokerNic 'Microsoft.Network/networkInterfaces@2023-04-01' = [
  for i in range(1, numberofCBs): {
    name: '${prefixCBName}-${i}-nic'
    location: location
    properties: {
      ipConfigurations: [
        {
          name: 'ipconfig1'
          properties: {
            privateIPAllocationMethod: privateIPAllocationMethod
            subnet: {
              id: subnetId
            }
          }
        }
      ]
      enableIPForwarding: true
    }
  }
]

resource connectionBrokerVM 'Microsoft.Compute/virtualMachines@2023-07-01' = [
  for i in range(1, numberofCBs): {
    name: '${prefixCBName}-${i}'
    location: location
    properties: {
      osProfile: {
        computerName: '${prefixCBName}-${i}'
        adminUsername: localAdminUser
        adminPassword: localAdminPassword
        windowsConfiguration: {
          enableAutomaticUpdates: false
          patchSettings: {
            patchMode: 'Manual'
          }
        }
      }
      hardwareProfile: {
        vmSize: vmSkuCB
      }
      storageProfile: {
        imageReference: {
          publisher: 'MicrosoftWindowsServer'
          offer: 'WindowsServer'
          sku: vmOSCB
          version: 'latest'
        }
        osDisk: {
          createOption: 'FromImage'
        }
      }
      networkProfile: {
        networkInterfaces: [
          {
            properties: {
              primary: true
            }
            id: ConnectionBrokerNic[i - 1].id
          }
        ]
      }
    }
  }
]

resource connectionBrokerDJ 'Microsoft.Compute/virtualMachines/extensions@2023-09-01' = [
  for i in range(1, numberofCBs): {
    parent: connectionBrokerVM[i - 1]
    name: 'joindomain'
    location: location
    properties: {
      publisher: 'Microsoft.Compute'
      type: 'JsonADDomainExtension'
      typeHandlerVersion: '1.3'
      autoUpgradeMinorVersion: true
      settings: {
        name: domainName
        ouPath: ouPath
        user: domainJoinUserName
        restart: true
        options: domainJoinOptions
      }
      protectedSettings: {
        Password: domainJoinPassword
      }
    }
  }
]

resource connectionBrokerPrimaryRAS 'Microsoft.Compute/virtualMachines/extensions@2023-07-01' = [
  for i in range(1, 1): {
    parent: connectionBrokerVM[i - 1]
    name: 'connectionBrokerPrimaryRAS'
    tags: {
      displayName: 'PowerShell Extension'
    }
    location: location
    properties: {
      publisher: 'Microsoft.Compute'
      type: 'CustomScriptExtension'
      typeHandlerVersion: '1.8'
      autoUpgradeMinorVersion: true
      settings: {
        fileUris: ['${assetLocation}${connectionBrokerPrimaryScript}']
      }
      protectedSettings: {
        commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${connectionBrokerPrimaryScript} -domainJoinUserName ${domainJoinUserName} -domainJoinPassword ${domainJoinPassword} -domainName ${domainName} -numberofCBs ${numberofCBs} -numberofSGs ${numberofSGs} -prefixCBName ${prefixCBName} -prefixSGName ${prefixSGName} -RasAdminsGroupAD ${RasAdminsGroupAD} -downloadURLRAS ${downloadURLRAS}'
      }
    }
    dependsOn: [connectionBrokerDJ, connectionBrokerRAS, secureGatewayRAS]
  }
]

resource connectionBrokerRAS 'Microsoft.Compute/virtualMachines/extensions@2023-07-01' = [
  for i in range(1, numberofCBs - 1): {
    parent: connectionBrokerVM[i]
    name: 'connectionBrokerRAS'
    tags: {
      displayName: 'PowerShell Extension'
    }
    location: location
    properties: {
      publisher: 'Microsoft.Compute'
      type: 'CustomScriptExtension'
      typeHandlerVersion: '1.8'
      autoUpgradeMinorVersion: true
      settings: {
        fileUris: ['${assetLocation}${prereqScript}']
      }
      protectedSettings: {
        commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${prereqScript} -domainJoinUserName ${domainJoinUserName}'
      }
    }
    dependsOn: [connectionBrokerDJ]
  }
]

resource secureGatewayNic 'Microsoft.Network/networkInterfaces@2023-04-01' = [
  for i in range(1, numberofSGs): {
    name: '${prefixSGName}-${i}-nic'
    location: location
    properties: {
      ipConfigurations: [
        {
          name: 'ipconfig1'
          properties: {
            privateIPAllocationMethod: privateIPAllocationMethod
            subnet: {
              id: subnetId
            }
            loadBalancerBackendAddressPools: [
              {
                id: secureGatewaylb.properties.backendAddressPools[0].id
              }
            ]
          }
        }
      ]
      enableIPForwarding: true
      networkSecurityGroup: {
        id: secureGatewayNSG.id
      }
    }
  }
]

resource secureGatewayVM 'Microsoft.Compute/virtualMachines@2023-07-01' = [
  for i in range(1, numberofSGs): {
    name: '${prefixSGName}-${i}'
    location: location
    properties: {
      osProfile: {
        computerName: '${prefixSGName}-${i}'
        adminUsername: localAdminUser
        adminPassword: localAdminPassword
        windowsConfiguration: {
          enableAutomaticUpdates: false
          patchSettings: {
            patchMode: 'Manual'
          }
        }
      }
      hardwareProfile: {
        vmSize: vmSkuSG
      }
      storageProfile: {
        imageReference: {
          publisher: 'MicrosoftWindowsServer'
          offer: 'WindowsServer'
          sku: vmOSSG
          version: 'latest'
        }
        osDisk: {
          createOption: 'FromImage'
        }
      }
      networkProfile: {
        networkInterfaces: [
          {
            properties: {
              primary: true
            }
            id: secureGatewayNic[i - 1].id
          }
        ]
      }
    }
  }
]

resource secureGatewayDJ 'Microsoft.Compute/virtualMachines/extensions@2023-09-01' = [
  for i in range(1, numberofSGs): {
    parent: secureGatewayVM[i - 1]
    name: 'joindomain'
    location: location
    properties: {
      publisher: 'Microsoft.Compute'
      type: 'JsonADDomainExtension'
      typeHandlerVersion: '1.3'
      autoUpgradeMinorVersion: true
      settings: {
        name: domainName
        ouPath: ouPath
        user: domainJoinUserName
        restart: true
        options: domainJoinOptions
      }
      protectedSettings: {
        Password: domainJoinPassword
      }
    }
  }
]

resource secureGatewayRAS 'Microsoft.Compute/virtualMachines/extensions@2023-07-01' = [
  for i in range(1, numberofSGs): {
    parent: secureGatewayVM[i - 1]
    name: 'secureGatewayRAS'
    tags: {
      displayName: 'PowerShell Extension'
    }
    location: location
    properties: {
      publisher: 'Microsoft.Compute'
      type: 'CustomScriptExtension'
      typeHandlerVersion: '1.8'
      autoUpgradeMinorVersion: true
      settings: {
        fileUris: ['${assetLocation}${prereqScript}']
      }
      protectedSettings: {
        commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${prereqScript} -domainJoinUserName ${domainJoinUserName}'
      }
    }
    dependsOn: [secureGatewayDJ[i - 1], secureGatewaylb]
  }
]

resource secureGatewayLBPIP 'Microsoft.Network/publicIPAddresses@2023-06-01' = {
  name: '${lbName}-pip'
  location: location
  sku: {
    name: 'Standard'
  }
  properties: {
    publicIPAllocationMethod: 'Static'
  }
}

resource secureGatewaylb 'Microsoft.Network/loadBalancers@2023-06-01' = {
  name: lbName
  location: location
  sku: {
    name: lbSkuName
  }
  properties: {
    frontendIPConfigurations: [
      {
        name: '${lbName}-frontend'
        properties: {
          publicIPAddress: {
            id: secureGatewayLBPIP.id
          }
        }
      }
    ]
    backendAddressPools: [
      {
        name: '${lbName}-pool'
      }
    ]
    loadBalancingRules: [
      {
        name: '${lbName}-443-Rule'
        properties: {
          frontendIPConfiguration: {
            id: resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', lbName, '${lbName}-frontend')
          }
          backendAddressPool: {
            id: resourceId('Microsoft.Network/loadBalancers/backendAddressPools', lbName, '${lbName}-pool')
          }
          frontendPort: 443
          backendPort: 443
          enableFloatingIP: false
          idleTimeoutInMinutes: 15
          protocol: 'Tcp'
          enableTcpReset: true
          loadDistribution: 'Default'
          disableOutboundSnat: true
          probe: {
            id: resourceId('Microsoft.Network/loadBalancers/probes', lbName, '${lbName}-443-Rule-Probe')
          }
        }
      }
    ]
    outboundRules: [
      {
        name: 'Outbound-traffic'
        properties: {
          backendAddressPool: {
            id: resourceId('Microsoft.Network/loadBalancers/backendAddressPools', lbName, '${lbName}-pool')
          }
          frontendIPConfigurations: [
            {
              id: resourceId('Microsoft.Network/loadBalancers/frontendIPConfigurations', lbName, '${lbName}-frontend')
            }
          ]
          protocol: 'All'
        }
      }
    ]
    probes: [
      {
        name: '${lbName}-443-Rule-Probe'
        properties: {
          protocol: 'Tcp'
          port: 80
          intervalInSeconds: 5
          numberOfProbes: 2
        }
      }
    ]
  }
}

resource secureGatewayNSG 'Microsoft.Network/networkSecurityGroups@2023-09-01' = {
  name: '${lbName}-NSG'
  location: location
  tags: {}
  properties: {
    securityRules: [
      {
        name: 'AllowAnyHTTPSInbound'
        type: 'Microsoft.Network/networkSecurityGroups/securityRules'
        properties: {
          protocol: 'TCP'
          sourcePortRange: '*'
          destinationPortRange: '443'
          sourceAddressPrefix: '*'
          destinationAddressPrefix: '*'
          access: 'Allow'
          priority: 100
          direction: 'Inbound'
        }
      }
      {
        name: 'AllowAnyHTTPInbound'
        type: 'Microsoft.Network/networkSecurityGroups/securityRules'
        properties: {
          protocol: 'TCP'
          sourcePortRange: '*'
          destinationPortRange: '80'
          sourceAddressPrefix: '*'
          destinationAddressPrefix: '*'
          access: 'Allow'
          priority: 101
          direction: 'Inbound'
        }
      }
    ]
  }
}

resource managementServerNic 'Microsoft.Network/networkInterfaces@2023-04-01' = {
  name: '${msName}-nic'
  location: location
  properties: {
    ipConfigurations: [
      {
        name: 'ipconfig1'
        properties: {
          privateIPAllocationMethod: privateIPAllocationMethod
          subnet: {
            id: subnetId
          }
        }
      }
    ]
    enableIPForwarding: true
  }
}

resource managementServerVM 'Microsoft.Compute/virtualMachines@2023-07-01' = {
  name: msName
  location: location
  properties: {
    osProfile: {
      computerName: msName
      adminUsername: localAdminUser
      adminPassword: localAdminPassword
      windowsConfiguration: {
        enableAutomaticUpdates: false
        patchSettings: {
          patchMode: 'Manual'
        }
      }
    }
    hardwareProfile: {
      vmSize: vmSkuMS
    }
    storageProfile: {
      imageReference: {
        publisher: 'MicrosoftWindowsServer'
        offer: 'WindowsServer'
        sku: vmOSMS
        version: 'latest'
      }
      osDisk: {
        createOption: 'FromImage'
      }
    }
    networkProfile: {
      networkInterfaces: [
        {
          properties: {
            primary: true
          }
          id: managementServerNic.id
        }
      ]
    }
  }
}

resource managementServerDJ 'Microsoft.Compute/virtualMachines/extensions@2023-09-01' = {
  parent: managementServerVM
  name: 'joindomain'
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'JsonADDomainExtension'
    typeHandlerVersion: '1.3'
    autoUpgradeMinorVersion: true
    settings: {
      name: domainName
      ouPath: ouPath
      user: domainJoinUserName
      restart: true
      options: domainJoinOptions
    }
    protectedSettings: {
      Password: domainJoinPassword
    }
  }
}

resource managementServerRAS 'Microsoft.Compute/virtualMachines/extensions@2023-07-01' = {
  parent: managementServerVM
  name: 'managementServerRAS'
  tags: {
    displayName: 'PowerShell Extension'
  }
  location: location
  properties: {
    publisher: 'Microsoft.Compute'
    type: 'CustomScriptExtension'
    typeHandlerVersion: '1.8'
    autoUpgradeMinorVersion: true
    settings: {
      fileUris: ['${assetLocation}${configurationScriptRAS}', '${assetLocation}${registerScriptRAS}']
    }
    protectedSettings: {
      commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${configurationScriptRAS} -domainJoinUserName ${domainJoinUserName} -domainJoinPassword ${domainJoinPassword} -domainName ${domainName} -resourceID ${managementServerVM.id} -tenantID ${tenant().tenantId} -keyVaultName ${keyVaultName} -secretName ${take(domainJoinUserName, indexOf(domainJoinUserName, '@')) } -primaryConnectionBroker ${prefixCBName}-${1} -numberofCBs ${numberofCBs} -numberofSGs ${numberofSGs} -prefixCBName ${prefixCBName} -prefixSGName ${prefixSGName} -appPublisherName ${appPublisherName} -appProductName ${appProductName} -customerUsageAttributionID ${customerUsageAttributionID} -providerSelection ${providerSelection} -providerName ${providerName} -providerAppRegistrationName ${providerAppRegistrationName} -vnetId ${vnetId[vnetNewOrExisting]} -mgrID ${resourceGroup().id} -downloadURLRAS ${downloadURLRAS} -licenseType ${licenseType}'
    }
  }
  dependsOn: [managementServerDJ, connectionBrokerRAS, connectionBrokerPrimaryRAS, secureGatewayRAS]
}

resource keyvault 'Microsoft.KeyVault/vaults@2023-02-01' = {
  name: keyVaultName
  location: location
  properties: {
    tenantId: azureADTenantID
    sku: {
      family: 'A'
      name: 'standard'
    }
    accessPolicies: accessPolicies
    enabledForDeployment: enabledForDeployment
    enabledForDiskEncryption: enabledForDiskEncryption
    enabledForTemplateDeployment: enabledForTemplateDeployment
    softDeleteRetentionInDays: softDeleteRetentionInDays
    enableRbacAuthorization: enableRbacAuthorization
  }
}

resource secretla 'Microsoft.KeyVault/vaults/secrets@2023-02-01' = {
  parent: keyvault
  name: localAdminUser
  properties: {
    value: localAdminPassword
  }
}

resource secretdj 'Microsoft.KeyVault/vaults/secrets@2023-02-01' = {
  parent: keyvault
  name: take(domainJoinUserName, indexOf(domainJoinUserName, '@'))
  properties: {
    value: domainJoinPassword
  }
}
