---
title: Inserir e excluir slides em uma apresentação do PowerPoint
description: Saiba como inserir slides de uma apresentação em outra e como excluir slides.
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613697"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a>Inserir e excluir slides em uma apresentação do PowerPoint (visualização)

Um suplemento do PowerPoint pode inserir slides de uma apresentação na apresentação atual usando a biblioteca JavaScript específica do aplicativo do PowerPoint. Você pode controlar se os slides inseridos manterão a formatação da apresentação de origem ou a formatação da apresentação de destino. Você também pode excluir slides da apresentação.

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

As APIs de inserção de slides são usadas principalmente nos cenários de modelo de apresentação: há um pequeno número de apresentações conhecidas que servem como pools de slides que podem ser inseridos pelo suplemento. Nesse cenário, você ou o cliente deve criar e manter uma fonte de dados que correlaciona o critério de seleção (como títulos de slides ou imagens) com IDs de slide. As APIs também podem ser usadas em cenários em que o usuário pode inserir slides de qualquer apresentação arbitrária, mas nesse cenário, o usuário está efetivamente limitado à inserção de *todos* os slides da apresentação de origem. Confira [selecionar quais slides inserir](#selecting-which-slides-to-insert) para obter mais informações sobre isso.

Há duas etapas para inserir slides de uma apresentação em outra.

1. Converta o arquivo de apresentação de origem (. pptx) em uma cadeia de caracteres formatada em base64.
1. Use o `insertSlidesFromBase64` método para inserir um ou mais slides do arquivo Base64 na apresentação atual.

## <a name="convert-the-source-presentation-to-base64"></a>Converter a apresentação de origem em base64

Há várias maneiras de converter um arquivo em base64. Qual linguagem de programação e biblioteca você usa, e se deseja converter no lado do servidor do seu suplemento ou no lado do cliente é determinado pelo seu cenário. Normalmente, você fará a conversão em JavaScript no lado do cliente usando um objeto [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . O exemplo a seguir mostra essa prática.

1. Comece obtendo uma referência para o arquivo de origem do PowerPoint. Neste exemplo, usaremos um `<input>` controle do tipo `file` para solicitar que o usuário escolha um arquivo. Adicione a seguinte marcação à página do suplemento.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Essa marcação adiciona a interface do usuário na seguinte captura de tela à página:

    ![Captura de tela mostrando um controle de entrada de tipo de arquivo HTML precedido por uma frase educacional lendo "selecionar uma apresentação do PowerPoint a partir da qual inserir slides". O controle consiste em um botão rotulado "escolher arquivo" seguido da frase "nenhum arquivo escolhido".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Há muitas outras maneiras de obter um arquivo do PowerPoint. Por exemplo, se o arquivo estiver armazenado no OneDrive ou no SharePoint, você poderá usar o Microsoft Graph para baixá-lo. Para obter mais informações, consulte [trabalhar com arquivos no Microsoft Graph](/graph/api/resources/onedrive) e [acessar arquivos com o Microsoft Graph](/learn/modules/msgraph-access-file-data/).

2. Adicione o código a seguir ao JavaScript do suplemento para atribuir uma função ao evento do controle de entrada `change` . (Você cria a `storeFileAsBase64` função na próxima etapa.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Adicione o código a seguir. Observe o seguinte sobre esse código:

    - O `reader.readAsDataURL` método converte o arquivo em Base64 e o armazena na `reader.result` propriedade. Quando o método é concluído, ele aciona o `onload` manipulador de eventos.
    - O `onload` manipulador de eventos apara os metadados do arquivo codificado e armazena a cadeia de caracteres codificada em uma variável global.
    - A cadeia de caracteres codificada em base64 é armazenada globalmente porque será lida por outra função que você cria em uma etapa posterior.

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a>Inserir slides com insertSlidesFromBase64

O suplemento insere slides de outra apresentação do PowerPoint na apresentação atual com o método [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) . Veja a seguir um exemplo simples no qual todos os slides da apresentação de origem são inseridos no início da apresentação atual e os slides inseridos mantêm a formatação do arquivo de origem. Observe que `chosenFileBase64` é uma variável global que contém uma versão codificada em Base64 de um arquivo de apresentação do PowerPoint.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Você pode controlar alguns aspectos do resultado da inserção, incluindo onde os slides são inseridos e se eles obtêm a formatação de origem ou de destino, passando um objeto [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) como um segundo parâmetro para `insertSlidesFromBase64` . Apresentamos um exemplo a seguir. Sobre este código, observe:

- Há dois valores possíveis para a `formatting` Propriedade: "UseDestinationTheme" e "KeepSourceFormatting". Opcionalmente, você pode usar a `InsertSlideFormatting` Enumeração (por exemplo, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).
- A função inserirá os slides da apresentação de origem imediatamente após o slide especificado pela `targetSlideId` propriedade. O valor dessa propriedade é uma cadeia de caracteres de uma das três formas possíveis:*** nnn * #**, *#* * mmmmmmmmm * * * ou **_nnn_ #* mmmmmmmmm * * *, onde *nnn* é a ID do slide (geralmente 3 dígitos) e *mmmmmmmmm* é a ID de criação do slide (normalmente, 9 dígitos). Alguns exemplos são `267#763315295` , `267#` e `#763315295` .

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

Obviamente, você normalmente não saberá no momento da codificação a ID ou a identificação de criação do slide de destino. Mais comumente, um suplemento solicitará que os usuários selecionem o slide de destino. As etapas a seguir mostram como obter o ***nnn * #** ID do slide selecionado no momento e usá-lo como o slide de destino.

1. Crie uma função que obtém a ID do slide selecionado no momento usando o método [Office.context.document. getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript comuns. Apresentamos um exemplo a seguir. Observe que a chamada para `getSelectedDataAsync` é incorporada em uma função de retorno de promessa. Para obter mais informações sobre por quê e como fazer isso, consulte [Wrap Common-APIs in Promise-retornod Functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. Chame sua nova função dentro do [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função main e passe a ID que ela retorna (concatenada com o símbolo "#") como o valor da `targetSlideId` Propriedade do `InsertSlideOptions` parâmetro. Apresentamos um exemplo a seguir.

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a>Selecionar os slides a serem inseridos

Você também pode usar o parâmetro [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) para controlar quais slides da apresentação de origem são inseridos. Para fazer isso, atribua uma matriz das IDs de slide da apresentação de origem à `sourceSlideIds` propriedade. Veja a seguir um exemplo que insere quatro slides. Observe que cada cadeia de caracteres na matriz deve seguir um ou outro dos padrões usados para a `targetSlideId` propriedade.

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> Os slides serão inseridos na mesma ordem relativa em que aparecem na apresentação de origem, independentemente da ordem em que aparecem na matriz.

Não há nenhuma maneira prática que os usuários possam descobrir a ID ou a identificação de criação de um slide na apresentação de origem. Por esse motivo, você só pode usar a `sourceSlideIds` propriedade quando você conhece as IDs de origem no tempo de codificação ou seu suplemento pode recuperá-las no tempo de execução de alguma fonte de dados. Como os usuários não podem se espera que memorizar as IDs de slide, você também precisa de uma maneira de permitir que o usuário selecione slides, talvez por título ou por uma imagem, e, em seguida, correlacione cada título ou imagem com a ID do slide.

Da mesma forma, a `sourceSlideIds` propriedade é usada principalmente em cenários de modelo de apresentação: o suplemento foi projetado para funcionar com um conjunto específico de apresentações que servem como pools de slides que podem ser inseridos. Nesse cenário, você ou o cliente deve criar e manter uma fonte de dados que correlaciona um critério de seleção (como títulos ou imagens) com IDs de slide ou IDs de criação de slides que foi construída a partir do conjunto de possíveis apresentações de origem.

## <a name="delete-slides"></a>Excluir slides

Você pode excluir um slide obtendo uma referência ao objeto [Slide](/javascript/api/powerpoint/powerpoint.slide) que representa o slide e chamar o `Slide.delete` método. A seguir está um exemplo no qual o 4º slide é excluído.

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
