---
title: Fazer sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: 90084fad0e79ab8acdf59eaa305825737401c0c8
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782823"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="5986e-102">Fazer sideload de Suplementos do Office para teste usando o **comando sideload**</span><span class="sxs-lookup"><span data-stu-id="5986e-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="5986e-103">O método "npm run sideload" funciona apenas para suplementos do Excel, Word e PowerPoint executados no Windows e para projetos de suplementos criados com a ferramenta [**yo office** e](https://github.com/OfficeDev/generator-office) que têm um script `sideload` na seção `scripts` do arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="5986e-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins that run on Windows; and only for add-in projects that were created with the [**yo office** tool](https://github.com/OfficeDev/generator-office) and that have a `sideload` script in the `scripts` section of the package.json file.</span></span> <span data-ttu-id="5986e-104">(Projetos criados com versões mais antigas do **yo office** também não têm esse script.) Se o seu projeto foi criado com o Visual Studio ou não tem o script de sideload, você pode fazer o sideload dele no Windows com o método descrito em [Fazer o sideload de um suplemento do Office a partir de um compartilhamento de rede](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5986e-104">(Projects that were created with older versions of **yo office** also do not have this script.) If your project was created with Visual Studio or does not have the sideload script, you can sideload it on Windows with the method described in [Sideload an Office add-in from a network share](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
>
> <span data-ttu-id="5986e-105">Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="5986e-105">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>
> 
> - [<span data-ttu-id="5986e-106">Sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="5986e-106">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
> - [<span data-ttu-id="5986e-107">Sideload dos suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="5986e-107">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [<span data-ttu-id="5986e-108">Fazer sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="5986e-108">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. <span data-ttu-id="5986e-109">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="5986e-109">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="5986e-110">Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="5986e-110">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="5986e-111">Execute o seguinte comando para iniciar uma instância do servidor da Web local na porta 3000 para servir seu projeto de suplemento: "**npm run start**"</span><span class="sxs-lookup"><span data-stu-id="5986e-111">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="5986e-112">Abra um segundo prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="5986e-112">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="5986e-113">Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="5986e-113">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="5986e-114">Execute o seguinte comando para inicializar o aplicativo host (por exemplo, Excel, Word) e registre seu suplemento no aplicativo host: "**npm run sideload**"</span><span class="sxs-lookup"><span data-stu-id="5986e-114">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="5986e-115">Veja também</span><span class="sxs-lookup"><span data-stu-id="5986e-115">See also</span></span>

- [<span data-ttu-id="5986e-116">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="5986e-116">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="5986e-117">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="5986e-117">Publish your Office Add-in</span></span>](../publish/publish.md)