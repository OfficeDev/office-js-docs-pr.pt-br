---
title: Caixas de diálogo em Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 5874e93c79d019adbd7ca5827a5b5e22f742190a
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457667"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="a25d6-102">Caixas de diálogo em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a25d6-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="a25d6-p101">Caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Office ativo. Você pode usar caixas de diálogo para fornecer espaço adicional na tela para tarefas como páginas de entrada que não podem ser abertas diretamente em um painel de tarefas ou solicitações para confirmar uma ação executada por um usuário ou mostrar vídeos que podem ser muito pequenos se confinados a um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a25d6-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="a25d6-105">*Figura 1. Layout típico de uma caixa de diálogo*</span><span class="sxs-lookup"><span data-stu-id="a25d6-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Uma imagem de exemplo que exibe um layout típico de uma caixa de diálogo](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="a25d6-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="a25d6-107">Best practices</span></span>

|<span data-ttu-id="a25d6-108">**Faça**</span><span class="sxs-lookup"><span data-stu-id="a25d6-108">**Do**</span></span>|<span data-ttu-id="a25d6-109">**Não faça**</span><span class="sxs-lookup"><span data-stu-id="a25d6-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="a25d6-110">Inclua um título descritivo com o nome de suplemento, juntamente com a tarefa atual.</span><span class="sxs-lookup"><span data-stu-id="a25d6-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="a25d6-111">Não adicione o nome da sua empresa ao título.</span><span class="sxs-lookup"><span data-stu-id="a25d6-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="a25d6-112">Não abra uma caixa de diálogo, a menos que o cenário exija isso.</span><span class="sxs-lookup"><span data-stu-id="a25d6-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="a25d6-113">Implementação</span><span class="sxs-lookup"><span data-stu-id="a25d6-113">Implementation</span></span>

<span data-ttu-id="a25d6-114">Confira um exemplo que implementa uma caixa de diálogo em [Exemplo de API de caixa de diálogo de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="a25d6-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="a25d6-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="a25d6-115">See also</span></span>

- [<span data-ttu-id="a25d6-116">Recursos de desenvolvimento do GitHub</span><span class="sxs-lookup"><span data-stu-id="a25d6-116">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="a25d6-117">Objeto Dialog</span><span class="sxs-lookup"><span data-stu-id="a25d6-117">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)


