param localAdminUser string = 'rasdmin'
@secure()
param localAdminPassword string

param vnetNameCB string
param vnetNewOrExistingCB string
param vnetkAddressPrefixCB string
param vnetResourceGroupCB string
param subnetNameCB string
param subnetAddressPrefixCB string

param vnetNameSG string = vnetNameCB
param vnetNewOrExistingSG string = vnetNewOrExistingCB
param vnetkAddressPrefixSG string = vnetkAddressPrefixCB
param vnetResourceGroupSG string = vnetResourceGroupCB
param subnetNameSG string = subnetNameCB
param subnetAddressPrefixSG string = subnetAddressPrefixCB

param vnetNameMS string = vnetNameCB
param vnetNewOrExistingMS string = vnetNewOrExistingCB
param vnetkAddressPrefixMS string = vnetkAddressPrefixCB
param vnetResourceGroupMS string = vnetResourceGroupCB
param subnetNameMS string = subnetNameCB
param subnetAddressPrefixMS string = subnetAddressPrefixCB

param vnetNameSH string = vnetNameCB
param vnetNewOrExistingSH string = vnetNewOrExistingCB
param vnetkAddressPrefixSH string = vnetkAddressPrefixCB
param vnetResourceGroupSH string = vnetResourceGroupCB
param subnetNameSH string = subnetNameCB
param subnetAddressPrefixSH string = subnetAddressPrefixCB

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
param providerSelection string
param providerName string = 'Provider'
param providerAppRegistrationName string = 'ras-app'
param license string
param rasVersion string = 'evergreen'
param customURLRAS string = 'evergreen'
@secure()
param maU string
@secure()
param maP string

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
var downloadURLRAS = rasVersion == 'customURL'
  ? customURLRAS
  : 'https://download.parallels.com/ras/latest/RASInstaller.msi'
var assetLocation = 'https://raw.githubusercontent.com/Parallels/RAS-PowerShell/master/RAS-Azure-Marketplace/'
var configurationScriptRAS = 'RAS_Azure_MP_Install.ps1'
var registerScriptRAS = 'RAS_Azure_MP_Register.ps1'
var prereqScript = 'RAS_Azure_MP_prereq.ps1'
var connectionBrokerPrimaryScript = 'RAS_Azure_MP_Primary_CB.ps1'
var privateIPAllocationMethod = 'Dynamic'

var vnetIdCB = {
  new: virtualNetworkCB.id
  existing: resourceId(vnetResourceGroupCB, 'Microsoft.Network/virtualNetworks', vnetNameCB)
}
var subnetIdCB = '${vnetIdCB[vnetNewOrExistingCB]}/subnets/${subnetNameCB}'

var vnetIdSG = {
  new: virtualNetworkCB.id
  existing: resourceId(vnetResourceGroupSG, 'Microsoft.Network/virtualNetworks', vnetNameSG)
}
var subnetIdSG = '${vnetIdSG[vnetNewOrExistingSG]}/subnets/${subnetNameSG}'

var vnetIdMS = {
  new: virtualNetworkCB.id
  existing: resourceId(vnetResourceGroupMS, 'Microsoft.Network/virtualNetworks', vnetNameMS)
}
var subnetIdMS = '${vnetIdMS[vnetNewOrExistingMS]}/subnets/${subnetNameMS}'

var vnetIdSH = {
  new: virtualNetworkCB.id
  existing: resourceId(vnetResourceGroupSH, 'Microsoft.Network/virtualNetworks', vnetNameSH)
}

var domainJoinOptions = 3
var lbSkuName = 'Standard'
var localAdminPasswordSecretName = 'localAdminPassword'
var domainJoinPasswordSecretName = 'domainJoinPassword'
var licenseType = 1

resource virtualNetworkCB 'Microsoft.Network/virtualNetworks@2023-06-01' = if (vnetNewOrExistingCB == 'new') {
  name: vnetNameCB
  location: location
  properties: {
    addressSpace: {
      addressPrefixes: [vnetkAddressPrefixCB]
    }
    subnets: [
      {
        name: subnetNameCB
        properties: {
          addressPrefix: subnetAddressPrefixCB
        }
      }
    ]
  }
}

resource virtualNetworkSG 'Microsoft.Network/virtualNetworks@2023-06-01' = if (vnetNewOrExistingSG == 'new') {
  name: vnetNameSG
  location: location
  properties: {
    addressSpace: {
      addressPrefixes: [vnetkAddressPrefixSG]
    }
    subnets: [
      {
        name: subnetNameSG
        properties: {
          addressPrefix: subnetAddressPrefixSG
        }
      }
    ]
  }
}

resource virtualNetworkMS 'Microsoft.Network/virtualNetworks@2023-06-01' = if (vnetNewOrExistingMS == 'new') {
  name: vnetNameMS
  location: location
  properties: {
    addressSpace: {
      addressPrefixes: [vnetkAddressPrefixMS]
    }
    subnets: [
      {
        name: subnetNameMS
        properties: {
          addressPrefix: subnetAddressPrefixMS
        }
      }
    ]
  }
}

resource virtualNetworkSH 'Microsoft.Network/virtualNetworks@2023-06-01' = if (vnetNewOrExistingSH == 'new') {
  name: vnetNameSH
  location: location
  properties: {
    addressSpace: {
      addressPrefixes: [vnetkAddressPrefixSH]
    }
    subnets: [
      {
        name: subnetNameSH
        properties: {
          addressPrefix: subnetAddressPrefixSH
        }
      }
    ]
  }
}

resource ConnectionBrokerNic 'Microsoft.Network/networkInterfaces@2023-04-01' = [
  for i in range(1, numberofCBs): {
    name: '${prefixCBName}-${padLeft(i, 2, '0')}-nic'
    location: location
    properties: {
      ipConfigurations: [
        {
          name: 'ipconfig1'
          properties: {
            privateIPAllocationMethod: privateIPAllocationMethod
            subnet: {
              id: subnetIdCB
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
    name: '${prefixCBName}-${padLeft(i, 2, '0')}'
    location: location
    properties: {
      osProfile: {
        computerName: '${prefixCBName}-${padLeft(i, 2, '0')}'
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
        commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${connectionBrokerPrimaryScript} -domainJoinUserName ${domainJoinUserName} -domainJoinPassword ${domainJoinPassword} -domainName ${domainName} -numberofCBs ${numberofCBs} -numberofSGs ${numberofSGs} -prefixCBName ${prefixCBName} -prefixSGName ${prefixSGName} -RasAdminsGroupAD ${RasAdminsGroupAD} -downloadURLRAS ${downloadURLRAS} -license ${license} -maU ${maU} -maP ${maP}'
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
    name: '${prefixSGName}-${padLeft(i, 2, '0')}-nic'
    location: location
    properties: {
      ipConfigurations: [
        {
          name: 'ipconfig1'
          properties: {
            privateIPAllocationMethod: privateIPAllocationMethod
            subnet: {
              id: subnetIdSG
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
    name: '${prefixSGName}-${padLeft(i, 2, '0')}'
    location: location
    properties: {
      osProfile: {
        computerName: '${prefixSGName}-${padLeft(i, 2, '0')}'
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
            id: subnetIdMS
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
      commandToExecute: 'powershell.exe -ExecutionPolicy Unrestricted -File ${configurationScriptRAS} -domainJoinUserName ${domainJoinUserName} -domainJoinPassword ${domainJoinPassword} -domainName ${domainName} -resourceID ${managementServerVM.id} -tenantID ${tenant().tenantId} -keyVaultName ${keyVaultName} -secretName ${domainJoinPasswordSecretName} -primaryConnectionBroker ${prefixCBName}-01 -numberofCBs ${numberofCBs} -numberofSGs ${numberofSGs} -prefixCBName ${prefixCBName} -prefixSGName ${prefixSGName} -appPublisherName ${appPublisherName} -appProductName ${appProductName} -providerSelection ${providerSelection} -providerName ${providerName} -providerAppRegistrationName ${providerAppRegistrationName} -vnetId ${vnetIdSH[vnetNewOrExistingSH]} -mgrID ${resourceGroup().id} -downloadURLRAS ${downloadURLRAS} -licenseType ${licenseType} '
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
  name: localAdminPasswordSecretName
  properties: {
    value: localAdminPassword
  }
}

resource secretdj 'Microsoft.KeyVault/vaults/secrets@2023-02-01' = {
  parent: keyvault
  name: domainJoinPasswordSecretName
  properties: {
    value: domainJoinPassword
  }
}
