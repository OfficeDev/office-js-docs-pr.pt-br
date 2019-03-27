---
title: Realizar sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870111"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="45727-102">Realizar sideload de Suplementos do Office usando o **comando sideload**</span><span class="sxs-lookup"><span data-stu-id="45727-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="45727-103">O método "npm executar sideload" só funciona para Word, Excel e PowerPoint suplementos executados no Windows; e somente para projetos que foi criado com a ferramenta [ **yo office** ](https://github.com/OfficeDev/generator-office) e que têm um `sideload` script na `scripts` seção do arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="45727-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="45727-104">(Projetos que foram criados com versões anteriores do **yo office** também não tem esse script.) Se o projeto foi criado com o Visual Studio ou não tem o script sideload, você pode sideload no Windows com o método descrito [Sideload um suplemento do Office em um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="45727-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office Add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="45727-105">Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="45727-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="45727-106">Sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="45727-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="45727-107">Sideload suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="45727-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="45727-108">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="45727-108">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="45727-109">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="45727-109">Open a command prompt as an administrator.</span></span>

2. <span data-ttu-id="45727-110">Altere os diretórios na raiz da pasta em um projeto.</span><span class="sxs-lookup"><span data-stu-id="45727-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="45727-111">Execute o seguinte comando para iniciar uma instância do servidor local da web na porta 3000 para atender a seu projeto do suplemento: "**npm executar início**"</span><span class="sxs-lookup"><span data-stu-id="45727-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="45727-112">Abra um segundo prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="45727-112">Open a second command prompt as an administrator.</span></span>

5. <span data-ttu-id="45727-113">Altere os diretórios na raiz da pasta em um projeto.</span><span class="sxs-lookup"><span data-stu-id="45727-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="45727-114">Execute o seguinte comando para inicializar o aplicativo de host (por exemplo, o Excel, Word) e inscreva-se o suplemento no aplicativo do host: "**npm executar sideload**"</span><span class="sxs-lookup"><span data-stu-id="45727-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="45727-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="45727-115">See also</span></span>

- [<span data-ttu-id="45727-116">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="45727-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="45727-117">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="45727-117">Publish your Office Add-in</span></span>](../publish/publish.md)
