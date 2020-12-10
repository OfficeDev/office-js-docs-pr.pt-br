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
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a><span data-ttu-id="b6c74-103">Inserir e excluir slides em uma apresentação do PowerPoint (visualização)</span><span class="sxs-lookup"><span data-stu-id="b6c74-103">Insert and delete slides in a PowerPoint presentation (preview)</span></span>

<span data-ttu-id="b6c74-104">Um suplemento do PowerPoint pode inserir slides de uma apresentação na apresentação atual usando a biblioteca JavaScript específica do aplicativo do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b6c74-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="b6c74-105">Você pode controlar se os slides inseridos manterão a formatação da apresentação de origem ou a formatação da apresentação de destino.</span><span class="sxs-lookup"><span data-stu-id="b6c74-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="b6c74-106">Você também pode excluir slides da apresentação.</span><span class="sxs-lookup"><span data-stu-id="b6c74-106">You can also delete slides from the presentation.</span></span>

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

<span data-ttu-id="b6c74-107">As APIs de inserção de slides são usadas principalmente nos cenários de modelo de apresentação: há um pequeno número de apresentações conhecidas que servem como pools de slides que podem ser inseridos pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="b6c74-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="b6c74-108">Nesse cenário, você ou o cliente deve criar e manter uma fonte de dados que correlaciona o critério de seleção (como títulos de slides ou imagens) com IDs de slide.</span><span class="sxs-lookup"><span data-stu-id="b6c74-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="b6c74-109">As APIs também podem ser usadas em cenários em que o usuário pode inserir slides de qualquer apresentação arbitrária, mas nesse cenário, o usuário está efetivamente limitado à inserção de *todos* os slides da apresentação de origem.</span><span class="sxs-lookup"><span data-stu-id="b6c74-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="b6c74-110">Confira [selecionar quais slides inserir](#selecting-which-slides-to-insert) para obter mais informações sobre isso.</span><span class="sxs-lookup"><span data-stu-id="b6c74-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="b6c74-111">Há duas etapas para inserir slides de uma apresentação em outra.</span><span class="sxs-lookup"><span data-stu-id="b6c74-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="b6c74-112">Converta o arquivo de apresentação de origem (. pptx) em uma cadeia de caracteres formatada em base64.</span><span class="sxs-lookup"><span data-stu-id="b6c74-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="b6c74-113">Use o `insertSlidesFromBase64` método para inserir um ou mais slides do arquivo Base64 na apresentação atual.</span><span class="sxs-lookup"><span data-stu-id="b6c74-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="b6c74-114">Converter a apresentação de origem em base64</span><span class="sxs-lookup"><span data-stu-id="b6c74-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="b6c74-115">Há várias maneiras de converter um arquivo em base64.</span><span class="sxs-lookup"><span data-stu-id="b6c74-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="b6c74-116">Qual linguagem de programação e biblioteca você usa, e se deseja converter no lado do servidor do seu suplemento ou no lado do cliente é determinado pelo seu cenário.</span><span class="sxs-lookup"><span data-stu-id="b6c74-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="b6c74-117">Normalmente, você fará a conversão em JavaScript no lado do cliente usando um objeto [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) .</span><span class="sxs-lookup"><span data-stu-id="b6c74-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="b6c74-118">O exemplo a seguir mostra essa prática.</span><span class="sxs-lookup"><span data-stu-id="b6c74-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="b6c74-119">Comece obtendo uma referência para o arquivo de origem do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b6c74-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="b6c74-120">Neste exemplo, usaremos um `<input>` controle do tipo `file` para solicitar que o usuário escolha um arquivo.</span><span class="sxs-lookup"><span data-stu-id="b6c74-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="b6c74-121">Adicione a seguinte marcação à página do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b6c74-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="b6c74-122">Essa marcação adiciona a interface do usuário na seguinte captura de tela à página:</span><span class="sxs-lookup"><span data-stu-id="b6c74-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![Captura de tela mostrando um controle de entrada de tipo de arquivo HTML precedido por uma frase educacional lendo "selecionar uma apresentação do PowerPoint a partir da qual inserir slides".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="b6c74-125">Há muitas outras maneiras de obter um arquivo do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b6c74-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="b6c74-126">Por exemplo, se o arquivo estiver armazenado no OneDrive ou no SharePoint, você poderá usar o Microsoft Graph para baixá-lo.</span><span class="sxs-lookup"><span data-stu-id="b6c74-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="b6c74-127">Para obter mais informações, consulte [trabalhar com arquivos no Microsoft Graph](/graph/api/resources/onedrive) e [acessar arquivos com o Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span><span class="sxs-lookup"><span data-stu-id="b6c74-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="b6c74-128">Adicione o código a seguir ao JavaScript do suplemento para atribuir uma função ao evento do controle de entrada `change` .</span><span class="sxs-lookup"><span data-stu-id="b6c74-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="b6c74-129">(Você cria a `storeFileAsBase64` função na próxima etapa.)</span><span class="sxs-lookup"><span data-stu-id="b6c74-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="b6c74-130">Adicione o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="b6c74-130">Add the following code.</span></span> <span data-ttu-id="b6c74-131">Observe o seguinte sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="b6c74-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="b6c74-132">O `reader.readAsDataURL` método converte o arquivo em Base64 e o armazena na `reader.result` propriedade.</span><span class="sxs-lookup"><span data-stu-id="b6c74-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="b6c74-133">Quando o método é concluído, ele aciona o `onload` manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b6c74-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="b6c74-134">O `onload` manipulador de eventos apara os metadados do arquivo codificado e armazena a cadeia de caracteres codificada em uma variável global.</span><span class="sxs-lookup"><span data-stu-id="b6c74-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="b6c74-135">A cadeia de caracteres codificada em base64 é armazenada globalmente porque será lida por outra função que você cria em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="b6c74-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="b6c74-136">Inserir slides com insertSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="b6c74-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="b6c74-137">O suplemento insere slides de outra apresentação do PowerPoint na apresentação atual com o método [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) .</span><span class="sxs-lookup"><span data-stu-id="b6c74-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="b6c74-138">Veja a seguir um exemplo simples no qual todos os slides da apresentação de origem são inseridos no início da apresentação atual e os slides inseridos mantêm a formatação do arquivo de origem.</span><span class="sxs-lookup"><span data-stu-id="b6c74-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="b6c74-139">Observe que `chosenFileBase64` é uma variável global que contém uma versão codificada em Base64 de um arquivo de apresentação do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b6c74-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="b6c74-140">Você pode controlar alguns aspectos do resultado da inserção, incluindo onde os slides são inseridos e se eles obtêm a formatação de origem ou de destino, passando um objeto [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) como um segundo parâmetro para `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="b6c74-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="b6c74-141">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="b6c74-141">The following is an example.</span></span> <span data-ttu-id="b6c74-142">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="b6c74-142">About this code, note:</span></span>

- <span data-ttu-id="b6c74-143">Há dois valores possíveis para a `formatting` Propriedade: "UseDestinationTheme" e "KeepSourceFormatting".</span><span class="sxs-lookup"><span data-stu-id="b6c74-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="b6c74-144">Opcionalmente, você pode usar a `InsertSlideFormatting` Enumeração (por exemplo, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).</span><span class="sxs-lookup"><span data-stu-id="b6c74-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="b6c74-145">A função inserirá os slides da apresentação de origem imediatamente após o slide especificado pela `targetSlideId` propriedade.</span><span class="sxs-lookup"><span data-stu-id="b6c74-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="b6c74-146">O valor dessa propriedade é uma cadeia de caracteres de uma das três formas possíveis:\*\*\* nnn \* #\*\*, *#* \* mmmmmmmmm \* \* \* ou \**_nnn_ #* mmmmmmmmm \* \* \*, onde *nnn* é a ID do slide (geralmente 3 dígitos) e *mmmmmmmmm* é a ID de criação do slide (normalmente, 9 dígitos).</span><span class="sxs-lookup"><span data-stu-id="b6c74-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="b6c74-147">Alguns exemplos são `267#763315295` , `267#` e `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="b6c74-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="b6c74-148">Obviamente, você normalmente não saberá no momento da codificação a ID ou a identificação de criação do slide de destino.</span><span class="sxs-lookup"><span data-stu-id="b6c74-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="b6c74-149">Mais comumente, um suplemento solicitará que os usuários selecionem o slide de destino.</span><span class="sxs-lookup"><span data-stu-id="b6c74-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="b6c74-150">As etapas a seguir mostram como obter o \***nnn \* #** ID do slide selecionado no momento e usá-lo como o slide de destino.</span><span class="sxs-lookup"><span data-stu-id="b6c74-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="b6c74-151">Crie uma função que obtém a ID do slide selecionado no momento usando o método [Office.context.document. getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript comuns.</span><span class="sxs-lookup"><span data-stu-id="b6c74-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="b6c74-152">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="b6c74-152">The following is an example.</span></span> <span data-ttu-id="b6c74-153">Observe que a chamada para `getSelectedDataAsync` é incorporada em uma função de retorno de promessa.</span><span class="sxs-lookup"><span data-stu-id="b6c74-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="b6c74-154">Para obter mais informações sobre por quê e como fazer isso, consulte [Wrap Common-APIs in Promise-retornod Functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="b6c74-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="b6c74-155">Chame sua nova função dentro do [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função main e passe a ID que ela retorna (concatenada com o símbolo "#") como o valor da `targetSlideId` Propriedade do `InsertSlideOptions` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="b6c74-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="b6c74-156">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="b6c74-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="b6c74-157">Selecionar os slides a serem inseridos</span><span class="sxs-lookup"><span data-stu-id="b6c74-157">Selecting which slides to insert</span></span>

<span data-ttu-id="b6c74-158">Você também pode usar o parâmetro [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) para controlar quais slides da apresentação de origem são inseridos.</span><span class="sxs-lookup"><span data-stu-id="b6c74-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="b6c74-159">Para fazer isso, atribua uma matriz das IDs de slide da apresentação de origem à `sourceSlideIds` propriedade.</span><span class="sxs-lookup"><span data-stu-id="b6c74-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="b6c74-160">Veja a seguir um exemplo que insere quatro slides.</span><span class="sxs-lookup"><span data-stu-id="b6c74-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="b6c74-161">Observe que cada cadeia de caracteres na matriz deve seguir um ou outro dos padrões usados para a `targetSlideId` propriedade.</span><span class="sxs-lookup"><span data-stu-id="b6c74-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="b6c74-162">Os slides serão inseridos na mesma ordem relativa em que aparecem na apresentação de origem, independentemente da ordem em que aparecem na matriz.</span><span class="sxs-lookup"><span data-stu-id="b6c74-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="b6c74-163">Não há nenhuma maneira prática que os usuários possam descobrir a ID ou a identificação de criação de um slide na apresentação de origem.</span><span class="sxs-lookup"><span data-stu-id="b6c74-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="b6c74-164">Por esse motivo, você só pode usar a `sourceSlideIds` propriedade quando você conhece as IDs de origem no tempo de codificação ou seu suplemento pode recuperá-las no tempo de execução de alguma fonte de dados.</span><span class="sxs-lookup"><span data-stu-id="b6c74-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="b6c74-165">Como os usuários não podem se espera que memorizar as IDs de slide, você também precisa de uma maneira de permitir que o usuário selecione slides, talvez por título ou por uma imagem, e, em seguida, correlacione cada título ou imagem com a ID do slide.</span><span class="sxs-lookup"><span data-stu-id="b6c74-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="b6c74-166">Da mesma forma, a `sourceSlideIds` propriedade é usada principalmente em cenários de modelo de apresentação: o suplemento foi projetado para funcionar com um conjunto específico de apresentações que servem como pools de slides que podem ser inseridos.</span><span class="sxs-lookup"><span data-stu-id="b6c74-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="b6c74-167">Nesse cenário, você ou o cliente deve criar e manter uma fonte de dados que correlaciona um critério de seleção (como títulos ou imagens) com IDs de slide ou IDs de criação de slides que foi construída a partir do conjunto de possíveis apresentações de origem.</span><span class="sxs-lookup"><span data-stu-id="b6c74-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="b6c74-168">Excluir slides</span><span class="sxs-lookup"><span data-stu-id="b6c74-168">Delete slides</span></span>

<span data-ttu-id="b6c74-169">Você pode excluir um slide obtendo uma referência ao objeto [Slide](/javascript/api/powerpoint/powerpoint.slide) que representa o slide e chamar o `Slide.delete` método.</span><span class="sxs-lookup"><span data-stu-id="b6c74-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="b6c74-170">A seguir está um exemplo no qual o 4º slide é excluído.</span><span class="sxs-lookup"><span data-stu-id="b6c74-170">The following is an example in which the 4th slide is deleted.</span></span>

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
