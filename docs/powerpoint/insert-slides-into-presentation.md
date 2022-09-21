---
title: Inserir slides em uma apresentação do PowerPoint
description: Saiba como inserir slides de uma apresentação em outra.
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: a31933de4272634394dc6c36aafa973c41265471
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857568"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Inserir slides em uma apresentação do PowerPoint

Um suplemento do PowerPoint pode inserir slides de uma apresentação na apresentação atual usando a biblioteca JavaScript específica do aplicativo do PowerPoint. Você pode controlar se os slides inseridos mantêm a formatação da apresentação de origem ou a formatação da apresentação de destino.

As APIs de inserção de slides são usadas principalmente em cenários de modelo de apresentação: há um pequeno número de apresentações conhecidas que servem como pools de slides que podem ser inseridos pelo suplemento. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona o critério de seleção (como títulos de slide ou imagens) com IDs de slide. As APIs também podem ser usadas em cenários em que o usuário pode inserir slides de qualquer apresentação arbitrária, mas nesse cenário o usuário está efetivamente  limitado a inserir todos os slides da apresentação de origem. Consulte [Selecionar quais slides inserir para](#selecting-which-slides-to-insert) obter mais informações sobre isso.

Há duas etapas para inserir slides de uma apresentação em outra.

1. Converta o arquivo de apresentação de origem (.pptx) em uma cadeia de caracteres formatada em base64.
1. Use o `insertSlidesFromBase64` método para inserir um ou mais slides do arquivo base64 na apresentação atual.

## <a name="convert-the-source-presentation-to-base64"></a>Converter a apresentação de origem em base64

Há muitas maneiras de converter um arquivo em base64. Qual linguagem de programação e biblioteca você usa e se a conversão no lado do servidor do suplemento ou do lado do cliente é determinada pelo seu cenário. Normalmente, você fará a conversão em JavaScript no lado do cliente usando um [objeto FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . O exemplo a seguir mostra essa prática.

1. Comece obtendo uma referência ao arquivo do PowerPoint de origem. Neste exemplo, usaremos um controle de `<input>` tipo para `file` solicitar que o usuário escolha um arquivo. Adicione a marcação a seguir à página do suplemento.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Essa marcação adiciona a interface do usuário na captura de tela a seguir à página.

    ![Captura de tela mostrando um controle de entrada de tipo de arquivo HTML precedido por uma frase instrucional que lê "Selecionar uma apresentação do PowerPoint da qual inserir slides". O controle consiste em um botão rotulado "Escolher arquivo" seguido da frase "Nenhum arquivo escolhido".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Há muitas outras maneiras de obter um arquivo do PowerPoint. Por exemplo, se o arquivo estiver armazenado no OneDrive ou no SharePoint, você poderá usar o Microsoft Graph para baixá-lo. Para obter mais informações, consulte [Trabalhando com arquivos no Microsoft Graph e](/graph/api/resources/onedrive) [acessar arquivos com o Microsoft Graph](/training/modules/msgraph-access-file-data/).

2. Adicione o código a seguir ao JavaScript do suplemento para atribuir uma função ao evento do controle de `change` entrada. (Você criará `storeFileAsBase64` a função na próxima etapa.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Adicione o código a seguir. Observe o seguinte sobre este código.

    - O `reader.readAsDataURL` método converte o arquivo em base64 e o armazena na `reader.result` propriedade. Quando o método é concluído, ele dispara o manipulador `onload` de eventos.
    - O `onload` manipulador de eventos corta metadados do arquivo codificado e armazena a cadeia de caracteres codificada em uma variável global.
    - A cadeia de caracteres codificada em base64 é armazenada globalmente porque ela será lida por outra função que você criará em uma etapa posterior.

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

O suplemento insere slides de outra apresentação do PowerPoint na apresentação atual com o método [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) . A seguir está um exemplo simples no qual todos os slides da apresentação de origem são inseridos no início da apresentação atual e os slides inseridos mantêm a formatação do arquivo de origem. Observe que `chosenFileBase64` é uma variável global que contém uma versão codificada em base64 de um arquivo de apresentação do PowerPoint.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Você pode controlar alguns aspectos do resultado da inserção, incluindo onde os slides são inseridos e se eles obtêm a `insertSlidesFromBase64`formatação de origem ou destino, passando um objeto [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) como um segundo parâmetro para . Apresentamos um exemplo a seguir. Sobre este código, observe:

- Há dois valores possíveis para a `formatting` propriedade: "UseDestinationTheme" e "KeepSourceFormatting". Opcionalmente, você pode usar `InsertSlideFormatting` a enumeração, (por exemplo, `PowerPoint.InsertSlideFormatting.useDestinationTheme`).
- A função inserirá os slides da apresentação de origem imediatamente após o slide especificado pela `targetSlideId` propriedade. O valor dessa propriedade é uma cadeia de caracteres de uma das três formas possíveis: ***nnn*#**, **#* mmmmmmmmm*** ou **_nnn_#* mmmmmmmmm***, em que *nnn* é a ID do slide (normalmente 3 dígitos) e *mmmmmmmmm* é a ID de criação do slide (normalmente, 9 dígitos). Alguns exemplos são `267#763315295`, `267#`e `#763315295`.

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

É claro que você normalmente não saberá no momento da codificação a ID ou a ID de criação do slide de destino. Mais comumente, um suplemento solicitará que os usuários selecionem o slide de destino. As etapas a seguir mostram como obter a ID ***nnn*#** do slide selecionado no momento e usá-la como o slide de destino.

1. Crie uma função que obtém a ID do slide selecionado no momento usando o método [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) das APIs JavaScript comuns. Apresentamos um exemplo a seguir. Observe que a chamada é inserida `getSelectedDataAsync` em uma função de retorno de promessa. Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common-APIs em funções de retorno de promessa](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

Você também pode usar o [parâmetro InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) para controlar quais slides da apresentação de origem são inseridos. Faça isso atribuindo uma matriz das IDs de slide da apresentação de origem à `sourceSlideIds` propriedade. A seguir está um exemplo que insere quatro slides. Observe que cada cadeia de caracteres na matriz deve seguir um ou outro dos padrões usados para a `targetSlideId` propriedade.

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

Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide na apresentação de origem. Por esse motivo, `sourceSlideIds` você só pode usar a propriedade quando souber as IDs de origem no momento da codificação ou se o suplemento puder recuperá-las em runtime de alguma fonte de dados. Como não é esperado que os usuários memorizem as IDs de slide, você também precisa de uma maneira de permitir que o usuário selecione slides, talvez por título ou por uma imagem, e, em seguida, correlacionar cada título ou imagem com a ID do slide.

Da mesma forma, `sourceSlideIds` a propriedade é usada principalmente em cenários de modelo de apresentação: o suplemento foi projetado para funcionar com um conjunto específico de apresentações que servem como pools de slides que podem ser inseridos. Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como títulos ou imagens) com IDs de slide ou IDs de criação de slide que foram construídas a partir do conjunto de possíveis apresentações de origem.
