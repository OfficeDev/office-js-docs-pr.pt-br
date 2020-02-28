---
title: Obter ou definir o local de um compromisso em um suplemento.
description: Saiba como obter ou definir o local de um compromisso em um suplemento do Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: cc412da5dd64d8e908b86a81b847f6479dbd4a34
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324965"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="f944b-103">Obter ou definir o local ao compor um compromisso no Outlook</span><span class="sxs-lookup"><span data-stu-id="f944b-103">Get or set the location when composing an appointment in Outlook</span></span>

<span data-ttu-id="f944b-104">A API JavaScript do Office fornece propriedades e métodos para gerenciar o local de um compromisso que o usuário está redigindo.</span><span class="sxs-lookup"><span data-stu-id="f944b-104">The Office JavaScript API provides properties and methods to manage the location of an appointment that the user is composing.</span></span> <span data-ttu-id="f944b-105">No momento, há duas propriedades que fornecem o local de um compromisso:</span><span class="sxs-lookup"><span data-stu-id="f944b-105">Currently, there are two properties that provide an appointment's location:</span></span>

- <span data-ttu-id="f944b-106">[Item. Location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API básica que permite obter e definir o local.</span><span class="sxs-lookup"><span data-stu-id="f944b-106">[item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.</span></span>
- <span data-ttu-id="f944b-107">[Item. enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API avançada que permite obter e definir o local e inclui a especificação do tipo de [local](/javascript/api/outlook/office.mailboxenums.locationtype).</span><span class="sxs-lookup"><span data-stu-id="f944b-107">[item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype).</span></span> <span data-ttu-id="f944b-108">O tipo é `LocationType.Custom` se você definir o local usando `item.location`.</span><span class="sxs-lookup"><span data-stu-id="f944b-108">The type is `LocationType.Custom` if you set the location using `item.location`.</span></span>

<span data-ttu-id="f944b-109">A tabela a seguir lista as APIs de local e os modos (ou seja, redigir ou ler) onde estão disponíveis.</span><span class="sxs-lookup"><span data-stu-id="f944b-109">The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.</span></span>

| <span data-ttu-id="f944b-110">API</span><span class="sxs-lookup"><span data-stu-id="f944b-110">API</span></span> | <span data-ttu-id="f944b-111">Modos de compromisso aplicáveis</span><span class="sxs-lookup"><span data-stu-id="f944b-111">Applicable appointment modes</span></span> |
|---|---|
| [<span data-ttu-id="f944b-112">item. Location</span><span class="sxs-lookup"><span data-stu-id="f944b-112">item.location</span></span>](/javascript/api/outlook/office.appointmentread#location) | <span data-ttu-id="f944b-113">Participante/leitura</span><span class="sxs-lookup"><span data-stu-id="f944b-113">Attendee/Read</span></span> |
| [<span data-ttu-id="f944b-114">item. Location. getasync</span><span class="sxs-lookup"><span data-stu-id="f944b-114">item.location.getAsync</span></span>](/javascript/api/outlook/office.location#getasync-options--callback-) | <span data-ttu-id="f944b-115">Organizador/compor</span><span class="sxs-lookup"><span data-stu-id="f944b-115">Organizer/Compose</span></span> |
| [<span data-ttu-id="f944b-116">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="f944b-116">item.location.setAsync</span></span>](/javascript/api/outlook/office.location#setasync-location--options--callback-) | <span data-ttu-id="f944b-117">Organizador/compor</span><span class="sxs-lookup"><span data-stu-id="f944b-117">Organizer/Compose</span></span> |
| [<span data-ttu-id="f944b-118">item. enhancedLocation. getasync</span><span class="sxs-lookup"><span data-stu-id="f944b-118">item.enhancedLocation.getAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | <span data-ttu-id="f944b-119">Organizador/compor,</span><span class="sxs-lookup"><span data-stu-id="f944b-119">Organizer/Compose,</span></span><br><span data-ttu-id="f944b-120">Participante/leitura</span><span class="sxs-lookup"><span data-stu-id="f944b-120">Attendee/Read</span></span> |
| [<span data-ttu-id="f944b-121">item. enhancedLocation. addasync</span><span class="sxs-lookup"><span data-stu-id="f944b-121">item.enhancedLocation.addAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | <span data-ttu-id="f944b-122">Organizador/compor</span><span class="sxs-lookup"><span data-stu-id="f944b-122">Organizer/Compose</span></span> |
| [<span data-ttu-id="f944b-123">item. enhancedLocation. removeAsync</span><span class="sxs-lookup"><span data-stu-id="f944b-123">item.enhancedLocation.removeAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | <span data-ttu-id="f944b-124">Organizador/compor</span><span class="sxs-lookup"><span data-stu-id="f944b-124">Organizer/Compose</span></span> |

<span data-ttu-id="f944b-125">Para usar os métodos que estão disponíveis somente para suplementos de composição, configure o manifesto do suplemento para ativar o suplemento no modo organizador/compor.</span><span class="sxs-lookup"><span data-stu-id="f944b-125">To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode.</span></span> <span data-ttu-id="f944b-126">Confira [criar suplementos do Outlook para formulários de redação](compose-scenario.md) para obter mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="f944b-126">See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.</span></span>

## <a name="use-the-enhancedlocation-api"></a><span data-ttu-id="f944b-127">Usar a `enhancedLocation` API</span><span class="sxs-lookup"><span data-stu-id="f944b-127">Use the `enhancedLocation` API</span></span>

<span data-ttu-id="f944b-128">Você pode usar a `enhancedLocation` API para obter e definir o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f944b-128">You can use the `enhancedLocation` API to get and set an appointment's location.</span></span> <span data-ttu-id="f944b-129">O campo local oferece suporte a vários locais e, para cada local, você pode definir o nome de exibição, o tipo e o endereço de email da sala de conferência (se aplicável).</span><span class="sxs-lookup"><span data-stu-id="f944b-129">The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable).</span></span> <span data-ttu-id="f944b-130">Consulte [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) para tipos de local com suporte.</span><span class="sxs-lookup"><span data-stu-id="f944b-130">See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.</span></span>

### <a name="add-location"></a><span data-ttu-id="f944b-131">Adicionar local</span><span class="sxs-lookup"><span data-stu-id="f944b-131">Add location</span></span>

<span data-ttu-id="f944b-132">O exemplo a seguir mostra como adicionar um local chamando [addasync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) em [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="f944b-132">The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a><span data-ttu-id="f944b-133">Obter local</span><span class="sxs-lookup"><span data-stu-id="f944b-133">Get location</span></span>

<span data-ttu-id="f944b-134">O exemplo a seguir mostra como obter o local chamando [getasync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) em [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="f944b-134">The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a><span data-ttu-id="f944b-135">Remover local</span><span class="sxs-lookup"><span data-stu-id="f944b-135">Remove location</span></span>

<span data-ttu-id="f944b-136">O exemplo a seguir mostra como remover o local chamando [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) na [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="f944b-136">The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a><span data-ttu-id="f944b-137">Usar a `location` API</span><span class="sxs-lookup"><span data-stu-id="f944b-137">Use the `location` API</span></span>

<span data-ttu-id="f944b-138">Você pode usar a `location` API para obter e definir o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f944b-138">You can use the `location` API to get and set an appointment's location.</span></span>

### <a name="get-the-location"></a><span data-ttu-id="f944b-139">Obter o local</span><span class="sxs-lookup"><span data-stu-id="f944b-139">Get the location</span></span>

<span data-ttu-id="f944b-140">Esta seção mostra um exemplo de código que obtém o local do compromisso que o usuário está compondo e o exibe.</span><span class="sxs-lookup"><span data-stu-id="f944b-140">This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.</span></span>

<span data-ttu-id="f944b-141">Para usar `item.location.getAsync`, forneça um método de retorno de chamada que verifica o status e o resultado da chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f944b-141">To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call.</span></span> <span data-ttu-id="f944b-142">Você pode fornecer os argumentos necessários para o método de retorno de chamada por meio do parâmetro opcional `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="f944b-142">You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter.</span></span> <span data-ttu-id="f944b-143">Você pode obter status, resultados e qualquer erro usando o parâmetro `asyncResult` de saída do retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f944b-143">You can obtain status, results, and any error using the output parameter `asyncResult` of the callback.</span></span> <span data-ttu-id="f944b-144">Se a chamada assíncrona for bem-sucedida, você poderá obter o local como uma cadeia de caracteres usando a propriedade [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="f944b-144">If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a><span data-ttu-id="f944b-145">Definir o local</span><span class="sxs-lookup"><span data-stu-id="f944b-145">Set the location</span></span>

<span data-ttu-id="f944b-146">Esta seção mostra um exemplo de código que define a localização do compromisso que o usuário está redigindo.</span><span class="sxs-lookup"><span data-stu-id="f944b-146">This section shows a code sample that sets the location of the appointment that the user is composing.</span></span>

<span data-ttu-id="f944b-147">Para usar `item.location.setAsync`, especifique uma cadeia de até 255 caracteres no parâmetro de dados.</span><span class="sxs-lookup"><span data-stu-id="f944b-147">To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter.</span></span> <span data-ttu-id="f944b-148">Opcionalmente, você pode fornecer um método de retorno de chamada e os argumentos para o método de retorno de chamada no parâmetro `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="f944b-148">Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter.</span></span> <span data-ttu-id="f944b-149">Você deve verificar o status, o resultado e qualquer mensagem de erro no `asyncResult` parâmetro de saída do retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f944b-149">You should check the status, result, and any error message in the `asyncResult` output parameter of the callback.</span></span> <span data-ttu-id="f944b-150">Se a chamada assíncrona for bem-sucedida, `setAsync` inserirá a cadeia de caracteres de local especificada como texto sem formatação, substituindo o local existente pelo item.</span><span class="sxs-lookup"><span data-stu-id="f944b-150">If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.</span></span>

> [!NOTE]
> <span data-ttu-id="f944b-151">Você pode definir vários locais usando um ponto-e-vírgula como separador (por exemplo, ' sala de conferência A; Sala de conferência B ').</span><span class="sxs-lookup"><span data-stu-id="f944b-151">You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a><span data-ttu-id="f944b-152">Confira também</span><span class="sxs-lookup"><span data-stu-id="f944b-152">See also</span></span>

- [<span data-ttu-id="f944b-153">Criar seu primeiro suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="f944b-153">Create your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="f944b-154">Programação assíncrona nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f944b-154">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
