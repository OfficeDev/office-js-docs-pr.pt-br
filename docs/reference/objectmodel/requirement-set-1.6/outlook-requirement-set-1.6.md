---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: Recursos e APIs que foram introduzidos para suplementos do Outlook e APIs JavaScript do Office como parte da API de caixa de correio 1,6.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: c1ce30ef1dd717a5d19ef9d71cf737e342cd660f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717632"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="a81bd-103">Conjunto de requisitos de API para suplementos do Outlook versão 1.6</span><span class="sxs-lookup"><span data-stu-id="a81bd-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="a81bd-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a81bd-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a81bd-105">Esta documentação se aplica a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="a81bd-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="a81bd-106">Novidades na versão 1.6</span><span class="sxs-lookup"><span data-stu-id="a81bd-106">What's new in 1.6?</span></span>

<span data-ttu-id="a81bd-107">O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="a81bd-107">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="a81bd-108">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="a81bd-108">It added the following features.</span></span>

- <span data-ttu-id="a81bd-109">Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a81bd-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="a81bd-110">Adicionada uma nova API para abrir um formulário de nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a81bd-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="a81bd-111">Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a81bd-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="a81bd-112">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="a81bd-112">Change log</span></span>

- <span data-ttu-id="a81bd-113">Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário.</span><span class="sxs-lookup"><span data-stu-id="a81bd-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="a81bd-114">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="a81bd-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="a81bd-115">Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a81bd-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="a81bd-116">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="a81bd-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="a81bd-117">Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): adiciona uma nova função que abre um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="a81bd-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="a81bd-118">Adicionado [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.</span><span class="sxs-lookup"><span data-stu-id="a81bd-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="a81bd-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="a81bd-119">See also</span></span>

- [<span data-ttu-id="a81bd-120">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="a81bd-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="a81bd-121">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="a81bd-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="a81bd-122">Introdução</span><span class="sxs-lookup"><span data-stu-id="a81bd-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a81bd-123">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="a81bd-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
