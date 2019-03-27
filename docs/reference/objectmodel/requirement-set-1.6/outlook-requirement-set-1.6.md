---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871323"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="91bc0-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.6</span><span class="sxs-lookup"><span data-stu-id="91bc0-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="91bc0-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="91bc0-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="91bc0-104">Esta documentação se aplica a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="91bc0-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="91bc0-105">Novidades na versão 1.6</span><span class="sxs-lookup"><span data-stu-id="91bc0-105">What's new in 1.6?</span></span>

<span data-ttu-id="91bc0-106">O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="91bc0-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="91bc0-107">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="91bc0-107">It added the following features.</span></span>

- <span data-ttu-id="91bc0-108">Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="91bc0-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="91bc0-109">Adicionada uma nova API para abrir um formulário de nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="91bc0-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="91bc0-110">Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="91bc0-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="91bc0-111">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="91bc0-111">Change log</span></span>

- <span data-ttu-id="91bc0-112">Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário.</span><span class="sxs-lookup"><span data-stu-id="91bc0-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="91bc0-113">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="91bc0-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="91bc0-114">Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="91bc0-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="91bc0-115">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="91bc0-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="91bc0-116">Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): adiciona uma nova função que abre um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="91bc0-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="91bc0-117">Adicionado [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.</span><span class="sxs-lookup"><span data-stu-id="91bc0-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="91bc0-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="91bc0-118">See also</span></span>

- [<span data-ttu-id="91bc0-119">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="91bc0-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="91bc0-120">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="91bc0-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="91bc0-121">Introdução</span><span class="sxs-lookup"><span data-stu-id="91bc0-121">Get started</span></span>](/outlook/add-ins/quick-start)
