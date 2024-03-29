---
title: Usar marcas personalizadas em apresentações, slides e formas no PowerPoint
description: Saiba como usar marcas para metadados personalizados sobre apresentações, slides e formas.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: a30beea56286437b1c69461534ca13912107cecf
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958899"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Usar marcas personalizadas para apresentações, slides e formas no PowerPoint

Um suplemento pode anexar metadados personalizados, na forma de pares chave-valor, chamados "marcas", a apresentações, slides específicos e formas específicas em um slide.

Há dois cenários principais para usar marcas:

- Quando aplicado a um slide ou uma forma, uma marca permite que o objeto seja categorizado para processamento em lotes. Por exemplo, suponha que uma apresentação tenha alguns slides que devem ser incluídos em apresentações para a região Leste, mas não para a região Oeste. Da mesma forma, há slides alternativos que devem ser mostrados somente para o Oeste. O suplemento pode criar uma marca `REGION` `East` com a chave e o valor e aplicá-la aos slides que só devem ser usados no Leste. O valor da marca é definido para `West` os slides que só devem ser mostrados para a região Oeste. Logo antes de uma apresentação para o Leste, um botão no suplemento executa o código que executa um loop em todos os slides verificando o valor da `REGION` marca. Os slides em que a região `West` está são excluídos. Em seguida, o usuário fecha o suplemento e inicia a apresentação de slides.
- Quando aplicada a uma apresentação, uma marca é efetivamente uma propriedade personalizada no documento de apresentação (semelhante a [uma CustomProperty](/javascript/api/word/word.customproperty) no Word).

## <a name="tag-slides-and-shapes"></a>Marcar slides e formas

Uma marca é um par chave-valor, em que o valor é sempre do tipo `string` e é representado por um [objeto Tag](/javascript/api/powerpoint/powerpoint.tag) . Cada tipo de objeto pai, como um objeto [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide) ou [Shape](/javascript/api/powerpoint/powerpoint.shape) , tem uma `tags` propriedade do tipo [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### <a name="add-update-and-delete-tags"></a>Adicionar, atualizar e excluir marcas

Para adicionar uma marca a um objeto, chame o [método TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) da propriedade do objeto `tags` pai. O código a seguir adiciona duas marcas ao primeiro slide de uma apresentação. Sobre este código, observe:

- O primeiro parâmetro do método `add` é a chave no par chave-valor.
- O segundo parâmetro é o valor.
- A chave está em letras maiúsculas. `add` Isso não é estritamente obrigatório para o método; no entanto, a chave é sempre armazenada pelo PowerPoint como maiúsculas e alguns métodos relacionados a marca exigem que a chave seja expressa em maiúsculas, portanto, recomendamos como prática recomendada que você sempre use *maiúsculas* em seu código para uma chave de marca.

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

O `add` método também é usado para atualizar uma marca. O código a seguir altera o valor da `PLANET` marca.

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

Para excluir uma marca, chame o `delete` método no `TagsCollection` objeto pai e passe a chave da marca como o parâmetro. Para obter um exemplo, consulte [Definir metadados personalizados na apresentação](#set-custom-metadata-on-the-presentation).

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Usar marcas para processar seletivamente slides e formas

Considere o seguinte cenário: a Contoso Consulting tem uma apresentação que eles mostram para todos os novos clientes. Mas alguns slides só devem ser mostrados aos clientes que pagaram pelo status "premium". Antes de mostrar a apresentação para clientes não premium, eles fazem uma cópia dela e excluem os slides que somente os clientes premium devem ver. Um suplemento permite que a Contoso marque quais slides são para clientes premium e exclua esses slides quando necessário. A lista a seguir descreve as principais etapas de codificação para criar essa funcionalidade.

1. Crie uma função que marca o slide selecionado no momento como destinado aos `Premium` clientes. Sobre este código, observe:

    - A `getSelectedSlideIndex` função é definida na próxima etapa. Ele retorna o índice baseado em 1 do slide selecionado no momento.
    - O valor retornado pela função `getSelectedSlideIndex` deve ser decrementado porque o método [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) é baseado em 0.

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

2. O código a seguir cria um método para obter o índice do slide selecionado. Sobre este código, observe:

    - Ele usa o [método Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) das APIs JavaScript comuns.
    - A chamada é `getSelectedDataAsync` inserida em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Encapsular APIs comuns em funções de retorno de promessa](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados. Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro (0º) slide, que é o único selecionado.
    - O `index` valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no painel de miniaturas da interface do usuário do PowerPoint.

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

3. O código a seguir cria uma função para excluir slides marcados para clientes premium. Sobre este código, observe:

    - Como as `key` propriedades `value` e as marcas serão lidas depois `context.sync`de , elas devem ser carregadas primeiro.

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

## <a name="set-custom-metadata-on-the-presentation"></a>Definir metadados personalizados na apresentação

Os suplementos também podem aplicar marcas à apresentação como um todo. Isso permite que você use marcas para metadados no nível do documento semelhantes a como a [classe CustomProperty](/javascript/api/word/word.customproperty)é usada no Word. Mas, ao contrário da classe do Word `CustomProperty` , o valor de uma marca do PowerPoint só pode ser do tipo `string`.

O código a seguir é um exemplo de como adicionar uma marca a uma apresentação. 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

O código a seguir é um exemplo de exclusão de uma marca de uma apresentação. Observe que a chave da marca é passada para o `delete` método do objeto `TagsCollection` pai.

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
