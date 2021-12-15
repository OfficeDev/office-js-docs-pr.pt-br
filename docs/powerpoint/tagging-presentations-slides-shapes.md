---
title: Use marcas personalizadas em apresentações, slides e formas em PowerPoint
description: Saiba como usar marcas para metadados personalizados sobre apresentações, slides e formas.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 03f1656919ed16b801e97623f7f69c9f4adfaac8
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514206"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Use marcas personalizadas para apresentações, slides e formas em PowerPoint

Um complemento pode anexar metadados personalizados, na forma de pares de valores-chave, chamados "marcas", a apresentações, slides específicos e formas específicas em um slide.

Há dois cenários principais para o uso de marcas:

- Quando aplicada a um slide ou uma forma, uma marca permite que o objeto seja categorizado para processamento em lotes. Por exemplo, suponha que uma apresentação tenha alguns slides que devem ser incluídos em apresentações para a região Leste, mas não para a região Oeste. Da mesma forma, há slides alternativos que devem ser mostrados somente para o O oeste. Seu complemento pode criar uma marca com a chave e o valor e aplicá-la aos slides que só devem ser `REGION` `East` usados no Leste. O valor da marca é definido como `West` para os slides que devem ser mostrados apenas para a região Oeste. Pouco antes de uma apresentação para o Leste, um botão no complemento executa o código que faz um loop por todos os slides verificando o valor da `REGION` marca. Slides onde a região `West` está são excluídos. Em seguida, o usuário fecha o complemento e inicia a apresentação de slides.
- Quando aplicada a uma apresentação, uma marca é efetivamente uma propriedade personalizada no documento de apresentação (semelhante a [uma CustomProperty](/javascript/api/word/word.customproperty) no Word).

## <a name="tag-slides-and-shapes"></a>Slides de marca e formas

Uma marca é um par de valores-chave, onde o valor é sempre do tipo e `string` é representado por um objeto [Tag.](/javascript/api/powerpoint/powerpoint.tag) Cada tipo de objeto pai, como um [objeto Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide)ou [Shape,](/javascript/api/powerpoint/powerpoint.shape) tem uma propriedade do `tags` tipo [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### <a name="add-update-and-delete-tags"></a>Adicionar, atualizar e excluir marcas

Para adicionar uma marca a um objeto, chame o [método TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) da propriedade do objeto `tags` pai. O código a seguir adiciona duas marcas ao primeiro slide de uma apresentação. Sobre este código, observe:

- O primeiro parâmetro do `add` método é a chave no par de valores-chave.
- O segundo parâmetro é o valor.
- A chave está em letras maiúsculas. Isso não é estritamente obrigatório para o método; no entanto, a chave é sempre armazenada pelo PowerPoint como maiúsculas, e alguns métodos relacionados a marca exigem que a chave seja expressa em maiúsculas, portanto, recomendamos como prática prática melhor que você sempre use maiúsculas em seu código para uma chave `add` de marca. 

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

Para excluir uma marca, chame o método em seu objeto pai e passe a chave `delete` da marca como o `TagsCollection` parâmetro. Para um exemplo, consulte [Definir metadados personalizados na apresentação](#set-custom-metadata-on-the-presentation).

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Usar marcas para processar seletivamente slides e formas

Considere o seguinte cenário: a Contoso Consulting tem uma apresentação que eles mostram para todos os novos clientes. Mas alguns slides só devem ser mostrados aos clientes que pagaram pelo status "premium". Antes de mostrar a apresentação para clientes não premium, eles fazem uma cópia dela e excluem os slides que apenas clientes premium devem ver. Um complemento permite que a Contoso marque quais slides são para clientes premium e exclua esses slides quando necessário. A lista a seguir descreve as principais etapas de codificação para criar essa funcionalidade.

1. Crie um método que marca o slide selecionado no momento como destinado aos `Premium` clientes. Sobre este código, observe:

    - A `getSelectedSlideIndex` função é definida na próxima etapa. Ele retorna o índice baseado em 1 do slide selecionado no momento.
    - O valor retornado pela função deve ser decrementado porque o método `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) é baseado em 0.

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

    - Ele usa o [método Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript Comuns.
    - A chamada para `getSelectedDataAsync` é inserida em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados. Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro slide (0th), que é o único selecionado.
    - O valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no painel PowerPoint `index` miniaturas da interface do usuário.

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

3. O código a seguir cria um método para excluir slides marcados para clientes premium. Sobre este código, observe:

    - Como as propriedades e das marcas serão lidas `key` depois do , eles devem ser `value` `context.sync` carregados primeiro.

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

Os complementos também podem aplicar marcas à apresentação como um todo. Isso permite que você use marcas para metadados no nível de documento semelhantes à forma como a [classe CustomProperty](/javascript/api/word/word.customproperty)é usada no Word. Mas, ao contrário da classe `CustomProperty` Word, o valor de uma marca PowerPoint só pode ser do tipo `string` .

O código a seguir é um exemplo de adição de uma marca a uma apresentação. 

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
