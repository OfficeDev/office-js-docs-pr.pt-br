---
title: Adicionar e excluir slides no PowerPoint
description: Saiba como adicionar e excluir slides e especificar o mestre e o layout de novos slides.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 915409c83e4eee2028a02f921e87065ee824bd7d
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514107"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Adicionar e excluir slides no PowerPoint

Um PowerPoint pode adicionar slides à apresentação e, opcionalmente, especificar qual slide mestre e qual layout do mestre é usado para o novo slide. O complemento também pode excluir slides.

As APIs para adicionar slides são usadas principalmente em cenários em que as IDs dos slides mestres e layouts da apresentação são conhecidas no momento da codificação ou podem ser encontradas em uma fonte de dados em tempo de execução. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona o critério de seleção (como nomes ou imagens de slides mestres e layouts) com as IDs dos slides mestres e layouts. As APIs também podem ser usadas em cenários em que o usuário pode inserir slides que usam o slide mestre padrão e o layout padrão do mestre e em cenários em que o usuário pode selecionar um slide existente e criar um novo com o mesmo slide mestre e layout (mas não o mesmo conteúdo). Confira [Selecionar o slide mestre e o layout a ser usado](#select-which-slide-master-and-layout-to-use) para obter mais informações sobre isso.

## <a name="add-a-slide-with-slidecollectionadd"></a>Adicionar um slide com SlideCollection.add

Adicione slides com o [método SlideCollection.add.](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) A seguir, um exemplo simples no qual um slide que usa o slide mestre padrão da apresentação e o primeiro layout desse mestre é adicionado. O método sempre adiciona novos slides ao final da apresentação. Apresentamos um exemplo a seguir.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Selecione qual slide mestre e layout usar

Use o [parâmetro AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) para controlar qual slide mestre é usado para o novo slide e qual layout dentro do mestre é usado. Apresentamos um exemplo a seguir. Sobre este código, observe:

- Você pode incluir as duas propriedades do `AddSlideOptions` objeto.
- Se ambas as propriedades são usadas, o layout especificado deve pertencer ao mestre especificado ou um erro é lançado.
- Se a propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o slide mestre padrão será usado e o deve ser um `masterId` `layoutId` layout desse slide mestre.
- O slide mestre padrão é o slide mestre usado pelo último slide da apresentação. (No caso incomum em que atualmente não há slides na apresentação, o slide mestre padrão é o primeiro slide mestre na apresentação.)
- Se a propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o primeiro layout do mestre especificado `layoutId` pelo `masterId` é usado.
- Ambas as propriedades são cadeias de caracteres de uma das três formas possíveis: ***nnnnnnnnnn*#**, * *#* mmmmmmmmm***, ou **_nnnnnnnnnn_ #* mmmmmmm****, onde *nnnnnnnnnn* é a ID do mestre ou layout (normalmente 10 dígitos) e *mmmmmmmmm* é a ID de criação do mestre ou layout (normalmente de 6 a 10 dígitos). Alguns exemplos são `2147483690#2908289500` , `2147483690#` e `#2908289500` .

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

Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide mestre ou layout. Por esse motivo, você só pode usar o parâmetro quando você conhece as IDs no momento da codificação ou seu complemento pode `AddSlideOptions` descobri-los em tempo de execução. Como não é esperado que os usuários memorizem as IDs, você também precisa de uma maneira de habilitar o usuário a selecionar slides, talvez por nome ou por uma imagem, e correlacionar cada título ou imagem com a ID do slide.

Portanto, o parâmetro é usado principalmente em cenários nos quais o complemento foi projetado para trabalhar com um conjunto específico de slides mestres e `AddSlideOptions` layouts cujas IDs são conhecidas. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como nomes ou imagens do slide mestre e layout) com as IDs ou IDs de criação correspondentes.

#### <a name="have-the-user-choose-a-matching-slide"></a>Fazer com que o usuário escolha um slide correspondente

Se o seu add-in puder ser usado em cenários em que o novo slide deve usar *a* mesma combinação de slide mestre e layout que é usado por um slide existente, seu complemento pode (1) solicitar que o usuário selecione um slide e (2) leia as IDs do slide mestre e layout. As etapas a seguir mostram como ler as IDs e adicionar um slide com um mestre e layout correspondentes.

1. Crie um método para obter o índice do slide selecionado. Apresentamos um exemplo a seguir. Sobre este código, observe:

    - Ele usa o [método Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript Comuns.
    - A chamada para `getSelectedDataAsync` é inserida em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados. Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro slide (0th), que é o único selecionado.
    - O valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no `index` painel de miniaturas.

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

2. Chame sua nova função dentro [do PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função principal que adiciona o slide. Apresentamos um exemplo a seguir.

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

## <a name="delete-slides"></a>Excluir slides

Exclua um slide recebendo uma referência ao [objeto Slide](/javascript/api/powerpoint/powerpoint.slide) que representa o slide e chame o `Slide.delete` método. A seguir, um exemplo no qual o 4º slide é excluído.

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
