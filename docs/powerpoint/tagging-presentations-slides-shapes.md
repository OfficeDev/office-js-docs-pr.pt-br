---
title: Usar marcas personalizadas em apresentações, slides e formas no PowerPoint
description: Saiba como usar marcas para metadados personalizados sobre apresentações, slides e formas.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fbb13e67da1f7962fc2c0b8d45689f259b015014
ms.sourcegitcommit: 58d394fa49308ecf93cd53f7d3fb6e316ff56209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/16/2021
ms.locfileid: "51876854"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a><span data-ttu-id="d2fd8-103">Usar marcas personalizadas para apresentações, slides e formas no PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d2fd8-103">Use custom tags for presentations, slides, and shapes in PowerPoint</span></span>

<span data-ttu-id="d2fd8-104">Um complemento pode anexar metadados personalizados, na forma de pares de valores-chave, chamados "marcas", a apresentações, slides específicos e formas específicas em um slide.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-104">An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d2fd8-105">As APIs para marcas estão em visualização.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-105">The APIs for tags are in preview.</span></span> <span data-ttu-id="d2fd8-106">Experimente-os em um ambiente de desenvolvimento ou teste, mas não os adicione a um complemento de produção.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-106">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>

<span data-ttu-id="d2fd8-107">Há dois cenários principais para o uso de marcas:</span><span class="sxs-lookup"><span data-stu-id="d2fd8-107">There are two main scenarios for using tags:</span></span>

- <span data-ttu-id="d2fd8-108">Quando aplicada a um slide ou uma forma, uma marca permite que o objeto seja categorizado para processamento em lotes.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-108">When applied to a slide or a shape, a tag enables the object to be categorized for batch processing.</span></span> <span data-ttu-id="d2fd8-109">Por exemplo, suponha que uma apresentação tenha alguns slides que devem ser incluídos em apresentações para a região Leste, mas não para a região Oeste.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-109">For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region.</span></span> <span data-ttu-id="d2fd8-110">Da mesma forma, há slides alternativos que devem ser mostrados somente para o O oeste.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-110">Similarly, there are alternative slides that should be shown only to the West.</span></span> <span data-ttu-id="d2fd8-111">Seu complemento pode criar uma marca com a chave e o valor e aplicá-la aos slides que só devem ser `REGION` `East` usados no Leste.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-111">Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East.</span></span> <span data-ttu-id="d2fd8-112">O valor da marca é definido como `West` para os slides que devem ser mostrados apenas para a região Oeste.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-112">The tag's value is set to `West` for the slides that should only be shown to the West region.</span></span> <span data-ttu-id="d2fd8-113">Pouco antes de uma apresentação para o Leste, um botão no complemento executa o código que faz um loop por todos os slides verificando o valor da `REGION` marca.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-113">Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag.</span></span> <span data-ttu-id="d2fd8-114">Slides onde a região `West` está são excluídos.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-114">Slides where the region is `West` are deleted.</span></span> <span data-ttu-id="d2fd8-115">Em seguida, o usuário fecha o complemento e inicia a apresentação de slides.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-115">The user then closes the add-in and starts the slide show.</span></span>
- <span data-ttu-id="d2fd8-116">Quando aplicada a uma apresentação, uma marca é efetivamente uma propriedade personalizada no documento de apresentação (semelhante a [uma CustomProperty](/javascript/api/word/word.customproperty) no Word).</span><span class="sxs-lookup"><span data-stu-id="d2fd8-116">When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).</span></span>

## <a name="tag-slides-and-shapes"></a><span data-ttu-id="d2fd8-117">Slides de marca e formas</span><span class="sxs-lookup"><span data-stu-id="d2fd8-117">Tag slides and shapes</span></span>

<span data-ttu-id="d2fd8-118">Uma marca é um par de valores-chave, onde o valor é sempre do tipo e `string` é representado por um objeto [Tag.](/javascript/api/powerpoint/powerpoint.tag)</span><span class="sxs-lookup"><span data-stu-id="d2fd8-118">A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object.</span></span> <span data-ttu-id="d2fd8-119">Cada tipo de objeto pai, como um [objeto Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide)ou [Shape,](/javascript/api/powerpoint/powerpoint.shape) tem uma propriedade do `tags` tipo [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span><span class="sxs-lookup"><span data-stu-id="d2fd8-119">Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span></span>

### <a name="add-update-and-delete-tags"></a><span data-ttu-id="d2fd8-120">Adicionar, atualizar e excluir marcas</span><span class="sxs-lookup"><span data-stu-id="d2fd8-120">Add, update, and delete tags</span></span>

<span data-ttu-id="d2fd8-121">Para adicionar uma marca a um objeto, chame o [método TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) da propriedade do objeto `tags` pai.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-121">To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) method of the parent object's `tags` property.</span></span> <span data-ttu-id="d2fd8-122">O código a seguir adiciona duas marcas ao primeiro slide de uma apresentação.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-122">The following code adds two tags to the first slide of a presentation.</span></span> <span data-ttu-id="d2fd8-123">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="d2fd8-123">About this code, note:</span></span>

- <span data-ttu-id="d2fd8-124">O primeiro parâmetro do `add` método é a chave no par de valores-chave.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-124">The first parameter of the `add` method is the key in the key-value pair.</span></span> 
- <span data-ttu-id="d2fd8-125">O segundo parâmetro é o valor.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-125">The second parameter is the value.</span></span>
- <span data-ttu-id="d2fd8-126">A chave está em letras maiúsculas.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-126">The key is in uppercase letters.</span></span> <span data-ttu-id="d2fd8-127">Isso não é estritamente obrigatório para o método; no entanto, a chave é sempre armazenada pelo PowerPoint como maiúsculas, e alguns métodos relacionados a marca exigem que a chave seja expressa em maiúsculas, portanto, recomendamos como prática de melhor prática que você sempre use maiúsculas em seu código para uma chave `add` de marca. </span><span class="sxs-lookup"><span data-stu-id="d2fd8-127">This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.</span></span>

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

<span data-ttu-id="d2fd8-128">O `add` método também é usado para atualizar uma marca.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-128">The `add` method is also used to update a tag.</span></span> <span data-ttu-id="d2fd8-129">O código a seguir altera o valor da `PLANET` marca.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-129">The following code changes the value of the `PLANET` tag.</span></span>

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

<span data-ttu-id="d2fd8-130">Para excluir uma marca, chame o método em seu objeto pai e passe a chave `delete` da marca como o `TagsCollection` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-130">To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter.</span></span> <span data-ttu-id="d2fd8-131">Para um exemplo, consulte [Definir metadados personalizados na apresentação](#set-custom-metadata-on-the-presentation).</span><span class="sxs-lookup"><span data-stu-id="d2fd8-131">For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).</span></span>

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a><span data-ttu-id="d2fd8-132">Usar marcas para processar seletivamente slides e formas</span><span class="sxs-lookup"><span data-stu-id="d2fd8-132">Use tags to selectively process slides and shapes</span></span>

<span data-ttu-id="d2fd8-133">Considere o seguinte cenário: a Contoso Consulting tem uma apresentação que eles mostram para todos os novos clientes.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-133">Consider the following scenario: Contoso Consulting has a presentation they show to all new customers.</span></span> <span data-ttu-id="d2fd8-134">Mas alguns slides só devem ser mostrados aos clientes que pagaram pelo status "premium".</span><span class="sxs-lookup"><span data-stu-id="d2fd8-134">But some slides should only be shown to customers that have paid for "premium" status.</span></span> <span data-ttu-id="d2fd8-135">Antes de mostrar a apresentação para clientes não premium, eles fazem uma cópia dela e excluem os slides que apenas clientes premium devem ver.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-135">Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see.</span></span> <span data-ttu-id="d2fd8-136">Um complemento permite que a Contoso marque quais slides são para clientes premium e exclua esses slides quando necessário.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-136">An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed.</span></span> <span data-ttu-id="d2fd8-137">A lista a seguir descreve as principais etapas de codificação para criar essa funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-137">The following list outlines the major coding steps to create this functionality.</span></span>

1. <span data-ttu-id="d2fd8-138">Crie um método que marca o slide selecionado no momento como destinado aos `Premium` clientes.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-138">Create a method that tags the currently selected slide as intended for `Premium` customers.</span></span> <span data-ttu-id="d2fd8-139">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="d2fd8-139">About this code, note:</span></span>

    - <span data-ttu-id="d2fd8-140">A `getSelectedSlideIndex` função é definida na próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-140">The `getSelectedSlideIndex` function is defined in the next step.</span></span> <span data-ttu-id="d2fd8-141">Ele retorna o índice baseado em 1 do slide selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-141">It returns the 1-based index of the currently selected slide.</span></span>
    - <span data-ttu-id="d2fd8-142">O valor retornado pela função deve ser decrementado porque o método `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) é baseado em 0.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-142">The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) method is 0-based.</span></span>

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. <span data-ttu-id="d2fd8-143">O código a seguir cria um método para obter o índice do slide selecionado.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-143">The following code creates a method to get the index of the selected slide.</span></span> <span data-ttu-id="d2fd8-144">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="d2fd8-144">About this code, note:</span></span>

    - <span data-ttu-id="d2fd8-145">Ele usa o [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript Comuns.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-145">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="d2fd8-146">A chamada para `getSelectedDataAsync` é inserida em uma função de retorno de promessa.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-146">The call to `getSelectedDataAsync` is embedded in a promise-returning function.</span></span> <span data-ttu-id="d2fd8-147">Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="d2fd8-147">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="d2fd8-148">`getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-148">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="d2fd8-149">Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro slide (0th), que é o único selecionado.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-149">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="d2fd8-150">O valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no painel miniaturas da interface do usuário do `index` PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-150">The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.</span></span>

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

3. <span data-ttu-id="d2fd8-151">O código a seguir cria um método para excluir slides marcados para clientes premium.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-151">The following code creates a method to delete slides that are tagged for premium customers.</span></span> <span data-ttu-id="d2fd8-152">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="d2fd8-152">About this code, note:</span></span>

    - <span data-ttu-id="d2fd8-153">Como as propriedades e das marcas serão lidas `key` depois do , eles devem ser `value` `context.sync` carregados primeiro.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-153">Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.</span></span>

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a><span data-ttu-id="d2fd8-154">Definir metadados personalizados na apresentação</span><span class="sxs-lookup"><span data-stu-id="d2fd8-154">Set custom metadata on the presentation</span></span>

<span data-ttu-id="d2fd8-155">Os complementos também podem aplicar marcas à apresentação como um todo.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-155">Add-ins can also apply tags to the presentation as a whole.</span></span> <span data-ttu-id="d2fd8-156">Isso permite que você use marcas para metadados no nível de documento semelhantes à forma como a [classe CustomProperty](/javascript/api/word/word.customproperty)é usada no Word.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-156">This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word.</span></span> <span data-ttu-id="d2fd8-157">Mas, ao contrário da classe `CustomProperty` Word, o valor de uma marca do PowerPoint só pode ser do tipo `string` .</span><span class="sxs-lookup"><span data-stu-id="d2fd8-157">But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.</span></span>

<span data-ttu-id="d2fd8-158">O código a seguir é um exemplo de adição de uma marca a uma apresentação.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-158">The following code is an example of adding a tag to a presentation.</span></span> 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

<span data-ttu-id="d2fd8-159">O código a seguir é um exemplo de exclusão de uma marca de uma apresentação.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-159">The following code is an example of deleting a tag from a presentation.</span></span> <span data-ttu-id="d2fd8-160">Observe que a chave da marca é passada para o `delete` método do objeto `TagsCollection` pai.</span><span class="sxs-lookup"><span data-stu-id="d2fd8-160">Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.</span></span>

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
