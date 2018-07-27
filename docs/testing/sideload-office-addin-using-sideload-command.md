---
title: Fazer sideload de Suplementos do Office usando o comando sideload
description: ''
ms.date: 07/24/2018
ms.openlocfilehash: e831a1dfbc31ecf06c8b2d78dc1e9a8a4c9dcf01
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21279357"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a><span data-ttu-id="96e6a-102">Fazer sideload de Suplementos do Office para teste usando o **comando sideload**</span><span class="sxs-lookup"><span data-stu-id="96e6a-102">Sideload Office Add-ins for testing using the **sideload command**</span></span>
 >[!NOTE]
><span data-ttu-id="96e6a-103">O método "npm run sideload" funciona apenas para suplementos do Excel, Word e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="96e6a-103">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

1. <span data-ttu-id="96e6a-104">Abra um prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="96e6a-104">Open a command prompt as administrator:</span></span>

2. <span data-ttu-id="96e6a-105">Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="96e6a-105">Change directories to the root of your add-in project folder.</span></span>

3. <span data-ttu-id="96e6a-106">Execute o seguinte comando para iniciar uma instância do servidor da Web local na porta 3000 para servir seu projeto de suplemento: "**npm run start**"</span><span class="sxs-lookup"><span data-stu-id="96e6a-106">Run the following command to start a local web server instance on port 3000 to serve up your add-in project: "**npm run start**"</span></span>

4. <span data-ttu-id="96e6a-107">Abrir um segundo prompt de comando como administrador.</span><span class="sxs-lookup"><span data-stu-id="96e6a-107">Open a command prompt as administrator:</span></span>

5. <span data-ttu-id="96e6a-108">Alterar os diretórios para a raiz da sua pasta de projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="96e6a-108">Change directories to the root of your add-in project folder.</span></span>

6. <span data-ttu-id="96e6a-109">Execute o seguinte comando para inicializar o aplicativo host (por exemplo, Excel, Word) e registre seu suplemento no aplicativo host: "**npm run sideload**"</span><span class="sxs-lookup"><span data-stu-id="96e6a-109">Run the following command to boot the host application (e.g. Excel, Word) and register your add-in in the host application: "**npm run sideload**"</span></span>

## <a name="see-also"></a><span data-ttu-id="96e6a-110">Confira também</span><span class="sxs-lookup"><span data-stu-id="96e6a-110">See also</span></span>

- [<span data-ttu-id="96e6a-111">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="96e6a-111">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="96e6a-112">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="96e6a-112">Publish your Office Add-in</span></span>](../publish/publish.md)