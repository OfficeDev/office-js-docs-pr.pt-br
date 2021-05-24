---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: cdb39eae387035f386a59b4640448b0bef25031e
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590992"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="a74b1-103">Conjunto de requisitos de API para suplementos do Outlook versão 1.6</span><span class="sxs-lookup"><span data-stu-id="a74b1-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="a74b1-104">O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="a74b1-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a74b1-105">Esta documentação se aplica a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="a74b1-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="a74b1-106">Novidades na versão 1.6</span><span class="sxs-lookup"><span data-stu-id="a74b1-106">What's new in 1.6?</span></span>

<span data-ttu-id="a74b1-107">O conjunto de requisitos 1.6 inclui todos os recursos do conjunto de requisitos [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="a74b1-107">Requirement set 1.6 includes all of the features of [requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="a74b1-108">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="a74b1-108">It added the following features.</span></span>

- <span data-ttu-id="a74b1-109">Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a74b1-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="a74b1-110">Adicionada uma nova API para abrir um formulário de nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a74b1-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="a74b1-111">Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a74b1-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="a74b1-112">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="a74b1-112">Change log</span></span>

- <span data-ttu-id="a74b1-113">Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário.</span><span class="sxs-lookup"><span data-stu-id="a74b1-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="a74b1-114">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="a74b1-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="a74b1-115">Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="a74b1-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="a74b1-116">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="a74b1-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="a74b1-117">Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): adiciona uma nova função que abre um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="a74b1-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="a74b1-118">Adicionado [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accounttype): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.</span><span class="sxs-lookup"><span data-stu-id="a74b1-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="a74b1-119">Confira também</span><span class="sxs-lookup"><span data-stu-id="a74b1-119">See also</span></span>

- [<span data-ttu-id="a74b1-120">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="a74b1-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="a74b1-121">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="a74b1-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="a74b1-122">Introdução</span><span class="sxs-lookup"><span data-stu-id="a74b1-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="a74b1-123">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="a74b1-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
