---
title: Adicionar e excluir slides no PowerPoint
description: Saiba como adicionar e excluir slides e especificar o mestre e o layout de novos slides.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf22c18cf4089bab9091be3f4274f67974662a3
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958310"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Adicionar e excluir slides no PowerPoint

Um suplemento do PowerPoint pode adicionar slides à apresentação e, opcionalmente, especificar qual slide mestre e qual layout do mestre é usado para o novo slide. O suplemento também pode excluir slides.

As APIs para adicionar slides são usadas principalmente em cenários em que as IDs dos slides mestres e layouts na apresentação são conhecidas no momento da codificação ou podem ser encontradas em uma fonte de dados em runtime. Nesse cenário, você ou o cliente deve criar e manter uma fonte de dados que correlaciona o critério de seleção (como nomes ou imagens de slides mestres e layouts) com as IDs dos slides mestres e layouts. As APIs também podem ser usadas em cenários em que o usuário pode inserir slides que usam o slide mestre padrão e o layout padrão do mestre e em cenários em que o usuário pode selecionar um slide existente e criar um novo com o mesmo slide mestre e layout (mas não o mesmo conteúdo). Confira [Selecionar qual slide mestre e layout usar](#select-which-slide-master-and-layout-to-use) para obter mais informações sobre isso.

## <a name="add-a-slide-with-slidecollectionadd"></a>Adicionar um slide com SlideCollection.add

Adicione slides com o [método SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) . A seguir está um exemplo simples no qual um slide que usa o slide mestre padrão da apresentação e o primeiro layout desse mestre é adicionado. O método sempre adiciona novos slides ao final da apresentação. Apresentamos um exemplo a seguir.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Selecionar qual slide mestre e layout usar

Use o [parâmetro AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) para controlar qual slide mestre é usado para o novo slide e qual layout dentro do mestre é usado. Apresentamos um exemplo a seguir. Sobre este código, observe:

- Você pode incluir uma ou ambas as propriedades do `AddSlideOptions` objeto.
- Se ambas as propriedades forem usadas, o layout especificado deverá pertencer ao mestre especificado ou um erro será gerado.
- Se a `masterId` propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o slide `layoutId` mestre padrão será usado e deverá ser um layout desse slide mestre.
- O slide mestre padrão é o slide mestre usado pelo último slide da apresentação. (No caso incomum em que atualmente não há slides na apresentação, o slide mestre padrão é o primeiro slide mestre na apresentação.)
- Se a `layoutId` propriedade não estiver presente (ou seu valor for uma cadeia de caracteres vazia), o primeiro layout do mestre especificado pelo `masterId` usuário será usado.
- Ambas as propriedades são cadeias de caracteres de uma das três formas possíveis: ***nnnnnnnnnn*#**, **#* mmmmmmmmm*** ou **_nnnnnnnnnn_#* mmmmmmmmm***, em que *nnnnnnnnnn* é a ID do mestre ou layout (normalmente 10 dígitos) e *mmmmmmmmm* é a ID de criação do mestre ou layout (normalmente de 6 a 10 dígitos). Alguns exemplos são `2147483690#2908289500`, `2147483690#`e `#2908289500`.

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

Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide mestre ou layout. Por esse motivo, você `AddSlideOptions` só pode usar o parâmetro quando souber as IDs no momento da codificação ou se o suplemento puder descobri-las em runtime. Como não é esperado que os usuários memorizem as IDs, você também precisa de uma maneira de permitir que o usuário selecione slides, talvez por nome ou por uma imagem, e, em seguida, correlacionar cada título ou imagem com a ID do slide.

Da mesma forma, `AddSlideOptions` o parâmetro é usado principalmente em cenários em que o suplemento foi projetado para funcionar com um conjunto específico de slides mestres e layouts cujas IDs são conhecidas. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como slide mestre e nomes de layout ou imagens) com as IDs correspondentes ou IDs de criação.

#### <a name="have-the-user-choose-a-matching-slide"></a>Fazer com que o usuário escolha um slide correspondente

Se o suplemento puder ser usado em cenários em que o novo slide deve usar a mesma combinação de slide mestre e layout que é usado por *um slide existente* , o suplemento pode (1) solicitar que o usuário selecione um slide e (2) leia as IDs do slide mestre e do layout. As etapas a seguir mostram como ler as IDs e adicionar um slide com um mestre e layout correspondentes.

1. Crie uma função para obter o índice do slide selecionado. Apresentamos um exemplo a seguir. Sobre este código, observe:

    - Ele usa o [método Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) das APIs JavaScript comuns.
    - A chamada é `getSelectedDataAsync` inserida em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Encapsular APIs comuns em funções de retorno de promessa](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` retorna uma matriz porque vários slides podem ser selecionados. Nesse cenário, o usuário selecionou apenas um, portanto, o código obtém o primeiro (0º) slide, que é o único selecionado.
    - O `index` valor do slide é o valor baseado em 1 que o usuário vê ao lado do slide no painel de miniaturas.

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

2. Chame sua nova função dentro do [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função principal que adiciona o slide. Apresentamos um exemplo a seguir.

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

Exclua um slide obtendo uma referência ao [objeto Slide](/javascript/api/powerpoint/powerpoint.slide) que representa o slide e chame o `Slide.delete` método. A seguir está um exemplo no qual o 4º slide é excluído.

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
