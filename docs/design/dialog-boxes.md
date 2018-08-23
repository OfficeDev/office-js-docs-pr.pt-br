---
title: Caixas de diálogo em Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: af47fd338872d3ecfce06145783fcc9ff314f7bc
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/23/2018
ms.locfileid: "19437140"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="f0565-102">Caixas de diálogo em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f0565-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="f0565-p101">Caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Office ativo. Você pode usar caixas de diálogo para fornecer espaço adicional na tela para tarefas como páginas de entrada que não podem ser abertas diretamente em um painel de tarefas ou solicitações para confirmar uma ação executada por um usuário ou mostrar vídeos que podem ser muito pequenos se confinados a um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f0565-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="f0565-105">*Figura 1. Layout típico de uma caixa de diálogo*</span><span class="sxs-lookup"><span data-stu-id="f0565-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Uma imagem de exemplo que exibe um layout típico de uma caixa de diálogo](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="f0565-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="f0565-107">Best practices</span></span>

|<span data-ttu-id="f0565-108">**Faça**</span><span class="sxs-lookup"><span data-stu-id="f0565-108">**Do**</span></span>|<span data-ttu-id="f0565-109">**Não faça**</span><span class="sxs-lookup"><span data-stu-id="f0565-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="f0565-110">Inclua um título descritivo com o nome de suplemento, juntamente com a tarefa atual.</span><span class="sxs-lookup"><span data-stu-id="f0565-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="f0565-111">Não adicione o nome da sua empresa ao título.</span><span class="sxs-lookup"><span data-stu-id="f0565-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="f0565-112">Não abra uma caixa de diálogo, a menos que o cenário exija isso.</span><span class="sxs-lookup"><span data-stu-id="f0565-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="f0565-113">Implementação</span><span class="sxs-lookup"><span data-stu-id="f0565-113">Implementation</span></span>

<span data-ttu-id="f0565-114">Confira um exemplo que implementa uma caixa de diálogo em [Exemplo de API de caixa de diálogo de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="f0565-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="f0565-115">Veja também</span><span class="sxs-lookup"><span data-stu-id="f0565-115">See also</span></span>

- [<span data-ttu-id="f0565-116">Amostra de padrão da experiência do usuário</span><span class="sxs-lookup"><span data-stu-id="f0565-116">UX Pattern Sample</span></span>](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [<span data-ttu-id="f0565-117">Recursos de desenvolvimento do GitHub</span><span class="sxs-lookup"><span data-stu-id="f0565-117">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="f0565-118">Objeto Dialog</span><span class="sxs-lookup"><span data-stu-id="f0565-118">Dialog object</span></span>](https://dev.office.com/reference/add-ins/shared/officeui.dialog)


