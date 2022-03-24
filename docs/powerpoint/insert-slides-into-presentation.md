---
title: Inserir slides em uma apresentação PowerPoint apresentação
description: Saiba como inserir slides de uma apresentação em outra.
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: b08dd8bd82e5d4f4f86114630e9238b6c43b6ae7
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747025"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Inserir slides em uma apresentação PowerPoint apresentação

Um PowerPoint pode inserir slides de uma apresentação na apresentação atual usando PowerPoint biblioteca JavaScript específica do aplicativo. Você pode controlar se os slides inseridos mantêm a formatação da apresentação de origem ou a formatação da apresentação de destino.

As APIs de inserção de slides são usadas principalmente em cenários de modelo de apresentação: há um pequeno número de apresentações conhecidas que servem como pools de slides que podem ser inseridos pelo complemento. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona o critério de seleção (como títulos de slide ou imagens) com IDs de slide. As APIs também podem ser usadas em cenários em que o usuário pode inserir slides de qualquer apresentação arbitrária, mas nesse cenário o usuário está efetivamente  limitado a inserir todos os slides da apresentação de origem. Confira [Selecionar quais slides inserir para](#selecting-which-slides-to-insert) obter mais informações sobre isso.

Há duas etapas para inserir slides de uma apresentação em outra.

1. Converta o arquivo de apresentação de origem (.pptx) em uma cadeia de caracteres formatada com base64.
1. Use o `insertSlidesFromBase64` método para inserir um ou mais slides do arquivo base64 na apresentação atual.

## <a name="convert-the-source-presentation-to-base64"></a>Converter a apresentação de origem em base64

Há muitas maneiras de converter um arquivo em base64. Qual linguagem de programação e biblioteca você usa e se a conversão no lado do servidor do seu complemento ou do lado do cliente é determinada pelo seu cenário. Mais comumente, você fará a conversão em JavaScript no lado do cliente usando um [objeto FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . O exemplo a seguir mostra essa prática.

1. Comece recebendo uma referência para o arquivo de PowerPoint de origem. Neste exemplo, vamos usar um controle de `<input>` tipo `file` para solicitar que o usuário escolha um arquivo. Adicione a marcação a seguir à página do complemento.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Essa marcação adiciona a interface do usuário na captura de tela a seguir à página.

    ![Captura de tela mostrando um controle de entrada de tipo de arquivo HTML precedido por uma frase instrucional que lê "Selecione uma apresentação PowerPoint da qual inserir slides". O controle consiste em um botão rotulado "Escolher arquivo" seguido da frase "Nenhum arquivo escolhido".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Há muitas outras maneiras de obter um arquivo PowerPoint. Por exemplo, se o arquivo for armazenado em OneDrive ou SharePoint, você poderá usar o Microsoft Graph baixá-lo. Para obter mais informações, consulte [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).

2. Adicione o código a seguir ao JavaScript do complemento para atribuir uma função ao evento do controle de `change` entrada. (Crie a `storeFileAsBase64` função na próxima etapa.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Adicione o código a seguir. Observe o seguinte sobre este código.

    - O `reader.readAsDataURL` método converte o arquivo em base64 e o armazena na `reader.result` propriedade. Quando o método é concluído, ele dispara o manipulador `onload` de eventos.
    - O `onload` manipulador de eventos corta metadados do arquivo codificado e armazena a cadeia de caracteres codificada em uma variável global.
    - A cadeia de caracteres codificada com base64 é armazenada globalmente porque ela será lida por outra função que você criar em uma etapa posterior.

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

Seu complemento insere slides de outra apresentação PowerPoint na apresentação atual com o método [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)). A seguir, um exemplo simples no qual todos os slides da apresentação de origem são inseridos no início da apresentação atual e os slides inseridos mantêm a formatação do arquivo de origem. Observe que `chosenFileBase64` é uma variável global que contém uma versão codificada com base64 de um arquivo PowerPoint apresentação.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Você pode controlar alguns aspectos do resultado de inserção, incluindo onde os slides são inseridos e se eles conseguem a formatação de origem ou de destino, passando um [objeto InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) como um segundo parâmetro para `insertSlidesFromBase64`. Apresentamos um exemplo a seguir. Sobre este código, observe:

- Há dois valores possíveis para a `formatting` propriedade: "UseDestinationTheme" e "KeepSourceFormatting". Opcionalmente, você pode usar `InsertSlideFormatting` o número , (por exemplo, `PowerPoint.InsertSlideFormatting.useDestinationTheme`).
- A função inserirá os slides da apresentação de origem imediatamente após o slide especificado pela `targetSlideId` propriedade. O valor dessa propriedade é uma cadeia de caracteres de uma das três formas possíveis: ***nnn*#**, **#* mmmmmmmmm***, ou **_nnnmmmmmmm_#****, onde *nnn* é a ID do slide (normalmente 3 dígitos) e *mmmmmmmmm* é a ID de criação do slide (normalmente 9 dígitos). Alguns exemplos são `267#763315295`, `267#`e `#763315295`.

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

Obviamente, você normalmente não saberá no momento da codificação a ID ou a ID de criação do slide de destino. Mais comumente, um complemento solicitará que os usuários selecionem o slide de destino. As etapas a seguir mostram como obter a ID ***nnn*#** do slide selecionado no momento e usá-lo como o slide de destino.

1. Crie uma função que obtém a ID do slide selecionado no momento usando o método [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) das APIs JavaScript Comuns. Apresentamos um exemplo a seguir. Observe que a chamada para `getSelectedDataAsync` está inserida em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

1. Chame sua nova função dentro do [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função principal e passe a ID que ela retorna (concatenada com o símbolo "#") `targetSlideId` `InsertSlideOptions` como o valor da propriedade do parâmetro. Apresentamos um exemplo a seguir.

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

### <a name="selecting-which-slides-to-insert"></a>Selecionando quais slides inserir

Você também pode usar o [parâmetro InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) para controlar quais slides da apresentação de origem são inseridos. Você faz isso atribuindo uma matriz das IDs de slide da apresentação de origem à `sourceSlideIds` propriedade. A seguir, um exemplo que insere quatro slides. Observe que cada cadeia de caracteres na matriz deve seguir um ou outro dos padrões usados para a `targetSlideId` propriedade.

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
> Os slides serão inseridos na mesma ordem relativa na qual aparecem na apresentação de origem, independentemente da ordem na qual aparecem na matriz.

Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide na apresentação de origem. Por esse motivo, você `sourceSlideIds` só pode usar a propriedade quando você sabe as IDs de origem no momento da codificação ou seu complemento pode recuperá-las em tempo de execução de alguma fonte de dados. Como não é esperado que os usuários memorizem IDs de slide, você também precisa de uma maneira de habilitar o usuário a selecionar slides, talvez por título ou por uma imagem, e correlacionar cada título ou imagem com a ID do slide.

Assim, `sourceSlideIds` a propriedade é usada principalmente em cenários de modelo de apresentação: o complemento foi projetado para funcionar com um conjunto específico de apresentações que servem como pools de slides que podem ser inseridos. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como títulos ou imagens) com IDs de slide ou IDs de criação de slide que foram construídas a partir do conjunto de possíveis apresentações de origem.
