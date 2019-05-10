---
title: Realizar sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: 69d39c2736312653b5a362aefccd41629e6e3555
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33619074"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="65ac3-102">Realizar sideload de Suplementos do Office usando o comando sideload</span><span class="sxs-lookup"><span data-stu-id="65ac3-102">Sideload Office Add-ins for testing using the sideload command</span></span>
 
> [!NOTE]
> <span data-ttu-id="65ac3-103">A técnica de sideload descrita neste artigo é válida somente para:</span><span class="sxs-lookup"><span data-stu-id="65ac3-103">The sideloading technique described in this article is only valid for:</span></span>
> 
> - <span data-ttu-id="65ac3-104">Suplementos do Excel, Word e PowerPoint executados no Windows</span><span class="sxs-lookup"><span data-stu-id="65ac3-104">Excel, Word, and PowerPoint add-ins that run on Windows</span></span>
> 
> - <span data-ttu-id="65ac3-105">Os projetos de suplemento que foram criados com o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) e que possuem um script `sideload` na seção `scripts` do arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="65ac3-105">Add-in projects that were created with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="65ac3-106">(Projetos que foram criados com as versões anteriores do gerador Yeoman para Suplementos do Office não possuirão este script.)</span><span class="sxs-lookup"><span data-stu-id="65ac3-106">(Projects that were created with older versions of the Yeoman generator for Office Add-ins will not have this script.)</span></span>
 
<span data-ttu-id="65ac3-107">Para realizar o sideload no seu suplemento usando o script `sideload` que o gerador Yeoman para Suplementos do Office fornece, conclua as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="65ac3-107">To sideload your add-in by using the `sideload` script that the Yeoman generator for Office Add-ins provides, complete the following steps:</span></span>

1. <span data-ttu-id="65ac3-108">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="65ac3-108">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="65ac3-109">Altere os diretórios na raiz da pasta em um projeto.</span><span class="sxs-lookup"><span data-stu-id="65ac3-109">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="65ac3-110">Execute o seguinte comando para iniciar uma instância do servidor local da web na porta 3000 para atender ao seu projeto de suplemento: `npm run start`</span><span class="sxs-lookup"><span data-stu-id="65ac3-110">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "`npm run start`"</span></span>

4. <span data-ttu-id="65ac3-111">Abra um segundo prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="65ac3-111">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="65ac3-112">Altere os diretórios na raiz da pasta em um projeto.</span><span class="sxs-lookup"><span data-stu-id="65ac3-112">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="65ac3-113">Execute o seguinte comando para inicializar o aplicativo de host (por exemplo, o Excel ou o Word) e registrar o seu suplemento no aplicativo do host: `npm run sideload`</span><span class="sxs-lookup"><span data-stu-id="65ac3-113">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "`npm run sideload`"</span></span>

<span data-ttu-id="65ac3-114">Se o seu projeto de suplemento foi criado com o Visual Studio ou não possui o script sideload, você pode realizar o sideload no Windows usando o método descrito em [Realizar Sideload em um Suplemento do Office a partir de um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="65ac3-114">(Projects that were created with older versions of yo office also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="65ac3-115">Se você não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para obter informações sobre como realizar o sideload no seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="65ac3-115">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
 
- [<span data-ttu-id="65ac3-116">Sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="65ac3-116">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="65ac3-117">Sideload suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="65ac3-117">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="65ac3-118">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="65ac3-118">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

## <a name="see-also"></a><span data-ttu-id="65ac3-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="65ac3-119">See also</span></span>

- [<span data-ttu-id="65ac3-120">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="65ac3-120">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="65ac3-121">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="65ac3-121">Publish your Office Add-in</span></span>](../publish/publish.md)
