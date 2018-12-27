---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.6
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: e780cff1a4cfe0751fccc9192784d143ab9c483f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433675"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="7845b-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.6</span><span class="sxs-lookup"><span data-stu-id="7845b-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="7845b-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="7845b-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7845b-104">Esta documentação se aplica a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="7845b-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="7845b-105">Novidades na versão 1.6</span><span class="sxs-lookup"><span data-stu-id="7845b-105">What's new in 1.1?</span></span>

<span data-ttu-id="7845b-106">O conjunto de requisitos 1.6 inclui todos os recursos do [Conjunto de requisitos 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="7845b-106">The Preview Requirement set includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="7845b-107">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="7845b-107">It added the following features.</span></span>

- <span data-ttu-id="7845b-108">Adicionadas novas APIs para suplementos contextuais para obter a correspondência de entidade ou regex que o usuário selecionou para ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7845b-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="7845b-109">Adicionada uma nova API para abrir um formulário de nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="7845b-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="7845b-110">Adicionada a capacidade de o suplemento determinar o tipo de conta da caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="7845b-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="7845b-111">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="7845b-111">Change log</span></span>

- <span data-ttu-id="7845b-112">Adicionado [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities): adiciona uma nova função que obtém as entidades encontradas em uma correspondência realçada selecionada por um usuário.</span><span class="sxs-lookup"><span data-stu-id="7845b-112">[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities) - Added a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="7845b-113">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="7845b-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="7845b-114">Adicionado [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): adiciona uma nova função que retorna valores de cadeia de caracteres em uma correspondência realçada que corresponde às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="7845b-114">[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) - Added a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="7845b-115">As correspondências realçadas aplicam-se aos suplementos contextuais.</span><span class="sxs-lookup"><span data-stu-id="7845b-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="7845b-116">Adicionado [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): adiciona uma nova função que abre um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="7845b-116">[Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters) - Added a new function that opens a new message form.</span></span>
- <span data-ttu-id="7845b-117">Adicionado [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): adiciona um novo membro ao perfil de usuário, que indica o tipo de conta do usuário.</span><span class="sxs-lookup"><span data-stu-id="7845b-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="7845b-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="7845b-118">See also</span></span>

- [<span data-ttu-id="7845b-119">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="7845b-119">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="7845b-120">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="7845b-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="7845b-121">Introdução</span><span class="sxs-lookup"><span data-stu-id="7845b-121">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)