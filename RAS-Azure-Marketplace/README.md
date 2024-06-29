# Automated deployment of Parallels RAS on Azure

## Description

This repository shares the resources to allow anyone to deploy Parallels RAS in any Azure subscription. The outcome of this deployment is a full-scale deployment of Parallels RAS in any Azure subscription. The template allows you to configure the number of Secure Gateway and Connection Broker components. Furthermore, you can easily connect to an existing vnet, subnet, and Active Directory.

The Bicep template is based on a Bring Your Own License or 30-day Trials and can optionally also configure a provider to connect to Azure Virtual Desktop. 

> Step 1: Visit this documentation page to learn about the [prerequisites](https://docs.parallels.com/parallels-ras-azure-marketplace-deployment-19/introduction/before-you-start)

> Step 2: Download the [Bicep template](https://github.com/Parallels/RAS-PowerShell/blob/master/RAS-Azure-Marketplace/mainTemplate.bicep) and adjust it if needed]

> Step 3: Deploy the Bicep template to your Azure Subscription. For more information on parameters and values to provide to the Bicep Template, [follow this guidance](https://docs.parallels.com/parallels-ras-azure-marketplace-deployment-19/deployment).

The screenshot below shows the outcome of the deployment.<br>
<image src=./images/deployment_result.png width=60%>

You can also use our [Azure Marketplace Transactional offer](https://azuremarketplace.microsoft.com/en-us/marketplace/apps/parallels.parallelsrasprod?tab=Overview) which uses the same template.

## License 

The scripts are MIT-licensed, so you are free to use them in your commercial setting.
