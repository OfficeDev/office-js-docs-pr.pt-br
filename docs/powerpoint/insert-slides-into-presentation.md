---
title: Inserir slides em uma apresentação PowerPoint apresentação
description: Saiba como inserir slides de uma apresentação em outra.
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 9b106e8940e7b0f19678e0467d8e900ffecd9438
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348780"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="8081d-103">Inserir slides em uma apresentação PowerPoint apresentação</span><span class="sxs-lookup"><span data-stu-id="8081d-103">Insert slides in a PowerPoint presentation</span></span>

<span data-ttu-id="8081d-104">Um PowerPoint pode inserir slides de uma apresentação na apresentação atual usando PowerPoint biblioteca JavaScript específica do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="8081d-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="8081d-105">Você pode controlar se os slides inseridos mantêm a formatação da apresentação de origem ou a formatação da apresentação de destino.</span><span class="sxs-lookup"><span data-stu-id="8081d-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span>

<span data-ttu-id="8081d-106">As APIs de inserção de slides são usadas principalmente em cenários de modelo de apresentação: há um pequeno número de apresentações conhecidas que servem como pools de slides que podem ser inseridos pelo complemento.</span><span class="sxs-lookup"><span data-stu-id="8081d-106">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="8081d-107">Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona o critério de seleção (como títulos de slide ou imagens) com IDs de slide.</span><span class="sxs-lookup"><span data-stu-id="8081d-107">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="8081d-108">As APIs também podem ser usadas em cenários em que o usuário pode inserir slides de  qualquer apresentação arbitrária, mas nesse cenário o usuário está efetivamente limitado a inserir todos os slides da apresentação de origem.</span><span class="sxs-lookup"><span data-stu-id="8081d-108">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="8081d-109">Confira [Selecionar quais slides inserir para](#selecting-which-slides-to-insert) obter mais informações sobre isso.</span><span class="sxs-lookup"><span data-stu-id="8081d-109">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="8081d-110">Há duas etapas para inserir slides de uma apresentação em outra.</span><span class="sxs-lookup"><span data-stu-id="8081d-110">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="8081d-111">Converta o arquivo de apresentação de origem (.pptx) em uma cadeia de caracteres formatada com base64.</span><span class="sxs-lookup"><span data-stu-id="8081d-111">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="8081d-112">Use o `insertSlidesFromBase64` método para inserir um ou mais slides do arquivo base64 na apresentação atual.</span><span class="sxs-lookup"><span data-stu-id="8081d-112">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="8081d-113">Converter a apresentação de origem em base64</span><span class="sxs-lookup"><span data-stu-id="8081d-113">Convert the source presentation to base64</span></span>

<span data-ttu-id="8081d-114">Há muitas maneiras de converter um arquivo em base64.</span><span class="sxs-lookup"><span data-stu-id="8081d-114">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="8081d-115">Qual linguagem de programação e biblioteca você usa e se a conversão no lado do servidor do seu complemento ou do lado do cliente é determinada pelo seu cenário.</span><span class="sxs-lookup"><span data-stu-id="8081d-115">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="8081d-116">Mais comumente, você fará a conversão em JavaScript no lado do cliente usando um [objeto FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader)</span><span class="sxs-lookup"><span data-stu-id="8081d-116">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="8081d-117">O exemplo a seguir mostra essa prática.</span><span class="sxs-lookup"><span data-stu-id="8081d-117">The following example shows this practice.</span></span>

1. <span data-ttu-id="8081d-118">Comece recebendo uma referência para o arquivo de PowerPoint de origem.</span><span class="sxs-lookup"><span data-stu-id="8081d-118">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="8081d-119">Neste exemplo, vamos usar um controle de tipo para solicitar que o `<input>` `file` usuário escolha um arquivo.</span><span class="sxs-lookup"><span data-stu-id="8081d-119">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="8081d-120">Adicione a marcação a seguir à página do complemento.</span><span class="sxs-lookup"><span data-stu-id="8081d-120">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="8081d-121">Essa marcação adiciona a interface do usuário na captura de tela a seguir à página.</span><span class="sxs-lookup"><span data-stu-id="8081d-121">This markup adds the UI in the following screenshot to the page.</span></span>

    ![Captura de tela mostrando um controle de entrada de tipo de arquivo HTML precedido por uma frase instrucional que lê "Selecione uma apresentação PowerPoint da qual inserir slides".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="8081d-124">Há muitas outras maneiras de obter um arquivo PowerPoint arquivo.</span><span class="sxs-lookup"><span data-stu-id="8081d-124">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="8081d-125">Por exemplo, se o arquivo for armazenado em OneDrive ou SharePoint, você poderá usar o Microsoft Graph baixá-lo.</span><span class="sxs-lookup"><span data-stu-id="8081d-125">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="8081d-126">Para obter mais informações, consulte [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and Access Files with Microsoft [Graph](/learn/modules/msgraph-access-file-data/).</span><span class="sxs-lookup"><span data-stu-id="8081d-126">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="8081d-127">Adicione o código a seguir ao JavaScript do complemento para atribuir uma função ao evento do controle de `change` entrada.</span><span class="sxs-lookup"><span data-stu-id="8081d-127">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="8081d-128">(Crie a `storeFileAsBase64` função na próxima etapa.)</span><span class="sxs-lookup"><span data-stu-id="8081d-128">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="8081d-129">Adicione o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="8081d-129">Add the following code.</span></span> <span data-ttu-id="8081d-130">Observe o seguinte sobre este código.</span><span class="sxs-lookup"><span data-stu-id="8081d-130">Note the following about this code.</span></span>

    - <span data-ttu-id="8081d-131">O `reader.readAsDataURL` método converte o arquivo em base64 e o armazena na `reader.result` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8081d-131">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="8081d-132">Quando o método é concluído, ele dispara o `onload` manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="8081d-132">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="8081d-133">O manipulador de eventos corta metadados do arquivo codificado e armazena a cadeia `onload` de caracteres codificada em uma variável global.</span><span class="sxs-lookup"><span data-stu-id="8081d-133">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="8081d-134">A cadeia de caracteres codificada com base64 é armazenada globalmente porque ela será lida por outra função que você criar em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="8081d-134">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="8081d-135">Inserir slides com insertSlidesFromBase64</span><span class="sxs-lookup"><span data-stu-id="8081d-135">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="8081d-136">O seu complemento insere slides de outra apresentação PowerPoint na apresentação atual com o [método Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-)</span><span class="sxs-lookup"><span data-stu-id="8081d-136">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="8081d-137">A seguir, um exemplo simples no qual todos os slides da apresentação de origem são inseridos no início da apresentação atual e os slides inseridos mantêm a formatação do arquivo de origem.</span><span class="sxs-lookup"><span data-stu-id="8081d-137">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="8081d-138">Observe que `chosenFileBase64` é uma variável global que contém uma versão codificada com base64 de um arquivo PowerPoint apresentação.</span><span class="sxs-lookup"><span data-stu-id="8081d-138">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="8081d-139">Você pode controlar alguns aspectos do resultado de inserção, incluindo onde os slides são inseridos e se eles conseguem a formatação de origem ou de destino, passando um [objeto InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) como um segundo parâmetro para `insertSlidesFromBase64` .</span><span class="sxs-lookup"><span data-stu-id="8081d-139">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="8081d-140">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8081d-140">The following is an example.</span></span> <span data-ttu-id="8081d-141">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="8081d-141">About this code, note:</span></span>

- <span data-ttu-id="8081d-142">Há dois valores possíveis para a `formatting` propriedade: "UseDestinationTheme" e "KeepSourceFormatting".</span><span class="sxs-lookup"><span data-stu-id="8081d-142">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="8081d-143">Opcionalmente, você pode usar `InsertSlideFormatting` o número , (por exemplo, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).</span><span class="sxs-lookup"><span data-stu-id="8081d-143">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="8081d-144">A função inserirá os slides da apresentação de origem imediatamente após o slide especificado pela `targetSlideId` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8081d-144">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="8081d-145">O valor dessa propriedade é uma cadeia de caracteres de uma das três formas possíveis: ***nnn\*#\*\*, \* *#* mmmmmmmmm***, ou \**_nnn_ #* mmmmmmmmm\*\*\*, onde *nnn* é a ID do slide (normalmente 3 dígitos) e *mmmmmmmmm* é a ID de criação do slide (normalmente 9 dígitos).</span><span class="sxs-lookup"><span data-stu-id="8081d-145">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="8081d-146">Alguns exemplos são `267#763315295` , `267#` e `#763315295` .</span><span class="sxs-lookup"><span data-stu-id="8081d-146">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="8081d-147">Obviamente, você normalmente não saberá no momento da codificação a ID ou a ID de criação do slide de destino.</span><span class="sxs-lookup"><span data-stu-id="8081d-147">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="8081d-148">Mais comumente, um complemento solicitará que os usuários selecionem o slide de destino.</span><span class="sxs-lookup"><span data-stu-id="8081d-148">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="8081d-149">As etapas a seguir mostram como obter a ID \***nnn\*#** do slide selecionado no momento e usá-lo como o slide de destino.</span><span class="sxs-lookup"><span data-stu-id="8081d-149">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="8081d-150">Crie uma função que obtém a ID do slide selecionado no momento usando o método [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) das APIs JavaScript Comuns.</span><span class="sxs-lookup"><span data-stu-id="8081d-150">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="8081d-151">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8081d-151">The following is an example.</span></span> <span data-ttu-id="8081d-152">Observe que a chamada para `getSelectedDataAsync` está inserida em uma função de retorno de promessa.</span><span class="sxs-lookup"><span data-stu-id="8081d-152">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="8081d-153">Para obter mais informações sobre por que e como fazer isso, consulte [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span><span class="sxs-lookup"><span data-stu-id="8081d-153">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="8081d-154">Chame sua nova função dentro do [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) da função principal e passe a ID que ela retorna (concatenada com o símbolo "#") como o valor da propriedade do `targetSlideId` `InsertSlideOptions` parâmetro.</span><span class="sxs-lookup"><span data-stu-id="8081d-154">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="8081d-155">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="8081d-155">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="8081d-156">Selecionando quais slides inserir</span><span class="sxs-lookup"><span data-stu-id="8081d-156">Selecting which slides to insert</span></span>

<span data-ttu-id="8081d-157">Você também pode usar o [parâmetro InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) para controlar quais slides da apresentação de origem são inseridos.</span><span class="sxs-lookup"><span data-stu-id="8081d-157">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="8081d-158">Você faz isso atribuindo uma matriz das IDs de slide da apresentação de origem à `sourceSlideIds` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8081d-158">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="8081d-159">A seguir, um exemplo que insere quatro slides.</span><span class="sxs-lookup"><span data-stu-id="8081d-159">The following is an example that inserts four slides.</span></span> <span data-ttu-id="8081d-160">Observe que cada cadeia de caracteres na matriz deve seguir um ou outro dos padrões usados para a `targetSlideId` propriedade.</span><span class="sxs-lookup"><span data-stu-id="8081d-160">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="8081d-161">Os slides serão inseridos na mesma ordem relativa na qual aparecem na apresentação de origem, independentemente da ordem na qual aparecem na matriz.</span><span class="sxs-lookup"><span data-stu-id="8081d-161">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="8081d-162">Não há nenhuma maneira prática de os usuários descobrirem a ID ou a ID de criação de um slide na apresentação de origem.</span><span class="sxs-lookup"><span data-stu-id="8081d-162">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="8081d-163">Por esse motivo, você só pode usar a propriedade quando você sabe as IDs de origem no momento da codificação ou seu complemento pode recuperá-las em tempo de execução de alguma fonte `sourceSlideIds` de dados.</span><span class="sxs-lookup"><span data-stu-id="8081d-163">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="8081d-164">Como não é esperado que os usuários memorizem IDs de slide, você também precisa de uma maneira de habilitar o usuário a selecionar slides, talvez por título ou por uma imagem, e correlacionar cada título ou imagem com a ID do slide.</span><span class="sxs-lookup"><span data-stu-id="8081d-164">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="8081d-165">Assim, a propriedade é usada principalmente em cenários de modelo de apresentação: o complemento foi projetado para funcionar com um conjunto específico de apresentações que servem como pools de slides que podem ser `sourceSlideIds` inseridos.</span><span class="sxs-lookup"><span data-stu-id="8081d-165">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="8081d-166">Nesse cenário, você ou o cliente devem criar e manter uma fonte de dados que correlaciona um critério de seleção (como títulos ou imagens) com IDs de slide ou IDs de criação de slide que foram construídas a partir do conjunto de possíveis apresentações de origem.</span><span class="sxs-lookup"><span data-stu-id="8081d-166">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>
