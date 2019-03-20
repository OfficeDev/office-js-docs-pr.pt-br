---
title: Caixas de diálogo em Suplementos do Office
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 3638006c30515a1fcae93ccfdbd43e0e92005c37
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691101"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="7f477-102">Caixas de diálogo em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7f477-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="7f477-p101">Caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Office ativo. Você pode usar caixas de diálogo para fornecer espaço adicional na tela para tarefas como páginas de entrada que não podem ser abertas diretamente em um painel de tarefas ou solicitações para confirmar uma ação executada por um usuário ou mostrar vídeos que podem ser muito pequenos se confinados a um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="7f477-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="7f477-105">*Figura 1. Layout típico de uma caixa de diálogo*</span><span class="sxs-lookup"><span data-stu-id="7f477-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Uma imagem de exemplo que exibe um layout típico de uma caixa de diálogo](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="7f477-107">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="7f477-107">Best practices</span></span>

|<span data-ttu-id="7f477-108">**Faça**</span><span class="sxs-lookup"><span data-stu-id="7f477-108">**Do**</span></span>|<span data-ttu-id="7f477-109">**Não faça**</span><span class="sxs-lookup"><span data-stu-id="7f477-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="7f477-110">Inclua um título descritivo com o nome de suplemento, juntamente com a tarefa atual.</span><span class="sxs-lookup"><span data-stu-id="7f477-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="7f477-111">Não adicione o nome da sua empresa ao título.</span><span class="sxs-lookup"><span data-stu-id="7f477-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="7f477-112">Não abra uma caixa de diálogo, a menos que o cenário exija isso.</span><span class="sxs-lookup"><span data-stu-id="7f477-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="7f477-113">Implementação</span><span class="sxs-lookup"><span data-stu-id="7f477-113">Implementation</span></span>

<span data-ttu-id="7f477-114">Confira um exemplo que implementa uma caixa de diálogo em [Exemplo de API de caixa de diálogo de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) no GitHub.</span><span class="sxs-lookup"><span data-stu-id="7f477-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="7f477-115">Confira também</span><span class="sxs-lookup"><span data-stu-id="7f477-115">See also</span></span>

- [<span data-ttu-id="7f477-116">Objeto Dialog</span><span class="sxs-lookup"><span data-stu-id="7f477-116">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [<span data-ttu-id="7f477-117">Padrões de design da experiência do usuário para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7f477-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)


