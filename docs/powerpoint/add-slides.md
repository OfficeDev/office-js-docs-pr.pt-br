---
title: Adicionar e excluir slides no PowerPoint
description: Saiba como adicionar e excluir slides e especificar o mestre e o layout de novos slides.
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 5c1b9750acb905fd8e92484bb960c70ba39a7ca9
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613939"
---
# <a name="add-and-delete-slides-in-powerpoint-preview"></a><span data-ttu-id="2023d-103">Adicionar e excluir slides no PowerPoint (visualização)</span><span class="sxs-lookup"><span data-stu-id="2023d-103">Add and delete slides in PowerPoint (preview)</span></span>

<span data-ttu-id="2023d-104">Um complemento do PowerPoint pode adicionar slides à apresentação e, opcionalmente, especificar qual slide mestre e qual layout do mestre é usado para o novo slide.</span><span class="sxs-lookup"><span data-stu-id="2023d-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="2023d-105">O complemento também pode excluir slides.</span><span class="sxs-lookup"><span data-stu-id="2023d-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2023d-106">As APIs para adicionar slides estão na visualização.</span><span class="sxs-lookup"><span data-stu-id="2023d-106">The APIs for adding slides are in preview.</span></span> <span data-ttu-id="2023d-107">Experimente-os em um ambiente de desenvolvimento ou teste, mas não os adicione a um complemento de produção.</span><span class="sxs-lookup"><span data-stu-id="2023d-107">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span> <span data-ttu-id="2023d-108">A API para *excluir* slides foi lançada.</span><span class="sxs-lookup"><span data-stu-id="2023d-108">The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="2023d-109">As APIs para adicionar slides são usadas principalmente em cenários em que as IDs dos slides mestres e layouts da apresentação são conhecidas no momento da codificação ou podem ser encontradas em uma fonte de dados em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="2023d-109">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="2023d-110">Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona o critério de seleção (como nomes ou imagens de slides mestres e layouts) com as IDs dos slides mestres e layouts.</span><span class="sxs-lookup"><span data-stu-id="2023d-110">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="2023d-111">As APIs também podem ser usadas em cenários em que o usuário pode inserir slides que usam o slide mestre padrão e o layout padrão do mestre e em cenários em que o usuário pode selecionar um slide existente e criar um novo com o mesmo slide mestre e layout (mas não o mesmo conteúdo).</span><span class="sxs-lookup"><span data-stu-id="2023d-111">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="2023d-112">Confira [Selecionar o slide mestre e o layout a ser usado](#selecting-which-slide-master-and-layout-to-use) para obter mais informações sobre isso.</span><span class="sxs-lookup"><span data-stu-id="2023d-112">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd"></a><span data-ttu-id="2023d-113">Adicionar um slide com SlideCollection.add</span><span class="sxs-lookup"><span data-stu-id="2023d-113">Add a slide with SlideCollection.add</span></span>

<span data-ttu-id="2023d-114">Adicione slides com o [método SlideCollection.add.](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)</span><span class="sxs-lookup"><span data-stu-id="2023d-114">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="2023d-115">A seguir, um exemplo simples no qual um slide que usa o slide mestre padrão da apresentação e o primeiro layout desse mestre é adicionado.</span><span class="sxs-lookup"><span data-stu-id="2023d-115">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="2023d-116">O método sempre adiciona novos slides ao final da apresentação.</span><span class="sxs-lookup"><span data-stu-id="2023d-116">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="2023d-117">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="2023d-117">The following is an example:</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="2023d-118">Selecionando qual slide mestre e layout usar</span><span class="sxs-lookup"><span data-stu-id="2023d-118">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="2023d-119">Use o [parâmetro AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) para controlar qual slide mestre é usado para o novo slide e qual layout dentro do mestre é usado.</span><span class="sxs-lookup"><span data-stu-id="2023d-119">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="2023d-120">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="2023d-120">The following is an example.</span></span> <span data-ttu-id="2023d-121">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="2023d-121">Note the following about this code:</span></span>

- <span data-ttu-id="2023d-122">Você pode incluir as duas propriedades do `AddSlideOptions` objeto.</span><span class="sxs-lookup"><span data-stu-id="2023d-122">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="2023d-123">Se ambas as propriedades são usadas, o layout especificado deve pertencer ao mestre especificado ou um erro é lançado.</span><span class="sxs-lookup"><span data-stu-id="2023d-123">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="2023d-124">Se a propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o slide mestre padrão será usado e o deve ser um `masterId` `layoutId` layout desse slide mestre.</span><span class="sxs-lookup"><span data-stu-id="2023d-124">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="2023d-125">O slide mestre padrão é o slide mestre usado pelo último slide da apresentação.</span><span class="sxs-lookup"><span data-stu-id="2023d-125">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="2023d-126">(No caso incomum em que atualmente não há slides na apresentação, o slide mestre padrão é o primeiro slide mestre na apresentação.)</span><span class="sxs-lookup"><span data-stu-id="2023d-126">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="2023d-127">Se a propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o primeiro layout do mestre especificado `layoutId` pelo `masterId` é usado.</span><span class="sxs-lookup"><span data-stu-id="2023d-127">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="2023d-128">Ambas as propriedades são cadeias de caracteres de uma das três formas possíveis: ***nnnnnnnnnn\*#\*\*, \* *#* mmmmmmmmm***, ou \**_nnnnnnnnnn_ #* mmmmmmm\*\*\*\*, onde *nnnnnnnnnn* é a ID do mestre ou layout (normalmente 10 dígitos) e *mmmmmmmmm* é a ID de criação do mestre ou layout (normalmente de 6 a 10 dígitos).</span><span class="sxs-lookup"><span data-stu-id="2023d-128">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="2023d-129">Alguns exemplos são `2147483690#2908289500` , `2147483690#` e `#2908289500` .</span><span class="sxs-lookup"><span data-stu-id="2023d-129">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

<span data-ttu-id="2023d-130">Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide mestre ou layout.</span><span class="sxs-lookup"><span data-stu-id="2023d-130">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="2023d-131">Por esse motivo, você só pode usar o parâmetro quando você conhece as IDs no momento da codificação ou seu complemento pode `AddSlideOptions` descobri-los em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="2023d-131">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="2023d-132">Como não é esperado que os usuários memorizem as IDs, você também precisa de uma maneira de habilitar o usuário a selecionar slides, talvez por nome ou por uma imagem, e correlacionar cada título ou imagem com a ID do slide.</span><span class="sxs-lookup"><span data-stu-id="2023d-132">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="2023d-133">Portanto, o parâmetro é usado principalmente em cenários nos quais o complemento foi projetado para trabalhar com um conjunto específico de slides mestres e `AddSlideOptions` layouts cujas IDs são conhecidas.</span><span class="sxs-lookup"><span data-stu-id="2023d-133">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="2023d-134">Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como nomes ou imagens do slide mestre e layout) com as IDs ou IDs de criação correspondentes.</span><span class="sxs-lookup"><span data-stu-id="2023d-134">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="2023d-135">Fazer com que o usuário escolha um slide correspondente</span><span class="sxs-lookup"><span data-stu-id="2023d-135">Have the user choose a matching slide</span></span>

<span data-ttu-id="2023d-136">Se o seu add-in puder ser usado em cenários em que o novo slide deve usar *a* mesma combinação de slide mestre e layout que é usado por um slide existente, seu complemento pode (1) solicitar que o usuário selecione um slide e (2) leia as IDs do slide mestre e layout.</span><span class="sxs-lookup"><span data-stu-id="2023d-136">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="2023d-137">As etapas a seguir mostram como ler as IDs e adicionar um slide com um mestre e layout correspondentes.</span><span class="sxs-lookup"><span data-stu-id="2023d-137">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="2023d-138">Crie um método para obter o índice do slide selecionado.</span><span class="sxs-lookup"><span data-stu-id="2023d-138">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="2023d-139">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="2023d-139">The following is an example.</span></span> <span data-ttu-id="2023d-140">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="2023d-140">Note about this code:</span></span>

    - <span data-ttu-id="2023d-141">Ele usa o [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript Comuns.</span><span class="sxs-lookup"><span data-stu-id="2023d-141">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="2023d-142">A chamada para `getSelectedDataAsync` é inserida em uma função de retorno de promessa.</span><span class="sxs-lookup"><span data-stu-id="2023d-142">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="2023d-143">Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="2023d-143">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="2023d-144">`getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados.</span><span class="sxs-lookup"><span data-stu-id="2023d-144">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="2023d-145">Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro slide (0th), que é o único selecionado.</span><span class="sxs-lookup"><span data-stu-id="2023d-145">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="2023d-146">O valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no `index` painel de miniaturas.</span><span class="sxs-lookup"><span data-stu-id="2023d-146">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. <span data-ttu-id="2023d-147">Chame sua nova função dentro [do PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função principal que adiciona o slide.</span><span class="sxs-lookup"><span data-stu-id="2023d-147">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="2023d-148">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="2023d-148">The following is an example:</span></span>

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a><span data-ttu-id="2023d-149">Excluir slides</span><span class="sxs-lookup"><span data-stu-id="2023d-149">Delete slides</span></span>

<span data-ttu-id="2023d-150">Exclua um slide recebendo uma referência ao [objeto Slide](/javascript/api/powerpoint/powerpoint.slide) que representa o slide e chame o `Slide.delete` método.</span><span class="sxs-lookup"><span data-stu-id="2023d-150">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="2023d-151">Veja a seguir um exemplo no qual o 4º slide é excluído:</span><span class="sxs-lookup"><span data-stu-id="2023d-151">The following is an example in which the 4th slide is deleted:</span></span>

```javascript
async function deleteSlide() {
    await PowerPoint.run(async function(context) {

        // The slide index is zero-based. 
        const slide = context.presentation.slides.getItemAt(3);
        slide.delete();

        await context.sync();
    });
}
```
