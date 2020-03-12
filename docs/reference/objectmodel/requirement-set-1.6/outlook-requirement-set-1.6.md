---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: ''
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 759286abe16e5bc32d6fd9ba1bb0e344d9d0a91c
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597030"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="6d891-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.6</span><span class="sxs-lookup"><span data-stu-id="6d891-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="6d891-103">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6d891-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="6d891-104">Esta documentação se aplica a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="6d891-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="6d891-105">Novidades na versão 1.6</span><span class="sxs-lookup"><span data-stu-id="6d891-105">What's new in 1.6?</span></span>

<span data-ttu-id="6d891-106">O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="6d891-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="6d891-107">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="6d891-107">It added the following features.</span></span>

- <span data-ttu-id="6d891-108">Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6d891-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="6d891-109">Adicionada uma nova API para abrir um formulário de nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="6d891-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="6d891-110">Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="6d891-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="6d891-111">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="6d891-111">Change log</span></span>

- <span data-ttu-id="6d891-112">Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário.</span><span class="sxs-lookup"><span data-stu-id="6d891-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="6d891-113">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="6d891-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="6d891-114">Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="6d891-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="6d891-115">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="6d891-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="6d891-116">Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): adiciona uma nova função que abre um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="6d891-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="6d891-117">Adicionado [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.</span><span class="sxs-lookup"><span data-stu-id="6d891-117">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="6d891-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="6d891-118">See also</span></span>

- [<span data-ttu-id="6d891-119">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="6d891-119">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="6d891-120">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="6d891-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="6d891-121">Introdução</span><span class="sxs-lookup"><span data-stu-id="6d891-121">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="6d891-122">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="6d891-122">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
