---
title: Atalhos de teclado personalizados em Office de complementos
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao seu Office Add-in.
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: c419731eec5c4707b04dd1e1e07d62aa3b0458a8
ms.sourcegitcommit: ba4fb7087b9841d38bb46a99a63e88df49514a4d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/05/2021
ms.locfileid: "52779338"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a><span data-ttu-id="bd68f-103">Adicionar atalhos de teclado personalizados aos seus Office de usuário</span><span class="sxs-lookup"><span data-stu-id="bd68f-103">Add custom keyboard shortcuts to your Office Add-ins</span></span>

<span data-ttu-id="bd68f-104">Atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do seu complemento funcionem com mais eficiência.</span><span class="sxs-lookup"><span data-stu-id="bd68f-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently.</span></span> <span data-ttu-id="bd68f-105">Atalhos de teclado também melhoram a acessibilidade do complemento para usuários com deficiências, fornecendo uma alternativa ao mouse.</span><span class="sxs-lookup"><span data-stu-id="bd68f-105">Keyboard shortcuts also improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="bd68f-106">Para começar com uma versão de trabalho de um add-in com atalhos de teclado já habilitados, clone e execute o exemplo Excel [Atalhos de Teclado.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="bd68f-106">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="bd68f-107">Quando você estiver pronto para adicionar atalhos de teclado ao seu próprio complemento, continue com este artigo.</span><span class="sxs-lookup"><span data-stu-id="bd68f-107">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="bd68f-108">Há três etapas para adicionar atalhos de teclado a um complemento:</span><span class="sxs-lookup"><span data-stu-id="bd68f-108">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="bd68f-109">[Configure o manifesto do complemento](#configure-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="bd68f-109">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="bd68f-110">[Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e atalhos de teclado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-110">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="bd68f-111">[Adicione uma ou mais chamadas de tempo de](#create-a-mapping-of-actions-to-their-functions) execução da API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear uma função para cada ação.</span><span class="sxs-lookup"><span data-stu-id="bd68f-111">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="bd68f-112">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="bd68f-112">Configure the manifest</span></span>

<span data-ttu-id="bd68f-113">Há duas pequenas alterações no manifesto a fazer.</span><span class="sxs-lookup"><span data-stu-id="bd68f-113">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="bd68f-114">Um deles é habilitar o add-in para usar um tempo de execução compartilhado e o outro é apontar para um arquivo formatado JSON onde você definiu os atalhos do teclado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-114">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="bd68f-115">Configurar o add-in para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="bd68f-115">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="bd68f-116">A adição de atalhos personalizados de teclado exige que o seu complemento use o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-116">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="bd68f-117">Para obter mais informações, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="bd68f-117">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="bd68f-118">Vincular o arquivo de mapeamento ao manifesto</span><span class="sxs-lookup"><span data-stu-id="bd68f-118">Link the mapping file to the manifest</span></span>

<span data-ttu-id="bd68f-119">Imediatamente *abaixo* (não dentro) `<VersionOverrides>` do elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="bd68f-119">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="bd68f-120">De definir o atributo como a URL completa de um arquivo JSON em `Url` seu projeto que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="bd68f-120">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="bd68f-121">Criar ou editar o arquivo JSON de atalhos</span><span class="sxs-lookup"><span data-stu-id="bd68f-121">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="bd68f-122">Crie um arquivo JSON em seu projeto.</span><span class="sxs-lookup"><span data-stu-id="bd68f-122">Create a JSON file in your project.</span></span> <span data-ttu-id="bd68f-123">Certifique-se de que o caminho do arquivo corresponde ao local especificado para o `Url` atributo do [elemento ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="bd68f-123">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="bd68f-124">Este arquivo descreverá seus atalhos de teclado e as ações que eles invocarão.</span><span class="sxs-lookup"><span data-stu-id="bd68f-124">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="bd68f-125">Dentro do arquivo JSON, há duas matrizes.</span><span class="sxs-lookup"><span data-stu-id="bd68f-125">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="bd68f-126">A matriz de ações conterá objetos que definem as ações a serem invocadas e a matriz de atalhos conterá objetos que mapeiam combinações de teclas em ações.</span><span class="sxs-lookup"><span data-stu-id="bd68f-126">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="bd68f-127">Veja um exemplo:</span><span class="sxs-lookup"><span data-stu-id="bd68f-127">Here is an example:</span></span>

    ```json
    {
        "actions": [
            {
                "id": "SHOWTASKPANE",
                "type": "ExecuteFunction",
                "name": "Show task pane for add-in"
            },
            {
                "id": "HIDETASKPANE",
                "type": "ExecuteFunction",
                "name": "Hide task pane for add-in"
            }
        ],
        "shortcuts": [
            {
                "action": "SHOWTASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Up"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "Ctrl+Alt+Down"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="bd68f-128">Para obter mais informações sobre os objetos JSON, consulte [Construct the action objects](#construct-the-action-objects) and Construct the shortcut [objects](#construct-the-shortcut-objects).</span><span class="sxs-lookup"><span data-stu-id="bd68f-128">For more information about the JSON objects, see [Construct the action objects](#construct-the-action-objects) and [Construct the shortcut objects](#construct-the-shortcut-objects).</span></span> <span data-ttu-id="bd68f-129">O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="bd68f-129">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="bd68f-130">Você pode usar "CONTROL" no lugar de "Ctrl" ao longo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="bd68f-130">You can use "CONTROL" in place of "Ctrl" throughout this article.</span></span>

    <span data-ttu-id="bd68f-131">Em uma etapa posterior, as ações serão mapeadas para as funções que você escrever.</span><span class="sxs-lookup"><span data-stu-id="bd68f-131">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="bd68f-132">Neste exemplo, mais tarde você mapeará SHOWTASKPANE para uma função que chama o método e `Office.addin.showAsTaskpane` HIDETASKPANE para uma função que chama o `Office.addin.hide` método.</span><span class="sxs-lookup"><span data-stu-id="bd68f-132">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="bd68f-133">Criar um mapeamento de ações para suas funções</span><span class="sxs-lookup"><span data-stu-id="bd68f-133">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="bd68f-134">Em seu projeto, abra o arquivo JavaScript carregado pela sua página HTML no `<FunctionFile>` elemento.</span><span class="sxs-lookup"><span data-stu-id="bd68f-134">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="bd68f-135">No arquivo JavaScript, use a API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear cada ação especificada no arquivo JSON para uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="bd68f-135">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="bd68f-136">Adicione o JavaScript a seguir ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd68f-136">Add the following JavaScript to the file.</span></span> <span data-ttu-id="bd68f-137">Observe o seguinte sobre o código:</span><span class="sxs-lookup"><span data-stu-id="bd68f-137">Note the following about the code:</span></span>

    - <span data-ttu-id="bd68f-138">O primeiro parâmetro é uma das ações do arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="bd68f-138">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="bd68f-139">O segundo parâmetro é a função que é executado quando um usuário pressiona a combinação de teclas mapeada para a ação no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="bd68f-139">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="bd68f-140">Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="bd68f-140">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="bd68f-141">Para o corpo da função, use o [método Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) para abrir o painel de tarefas do complemento.</span><span class="sxs-lookup"><span data-stu-id="bd68f-141">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="bd68f-142">Quando terminar, o código deverá ter a seguinte aparência:</span><span class="sxs-lookup"><span data-stu-id="bd68f-142">When you are done, the code should look like the following:</span></span>

    ```javascript
    Office.actions.associate('SHOWTASKPANE', function () {
        return Office.addin.showAsTaskpane()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

1. <span data-ttu-id="bd68f-143">Adicione uma segunda chamada de função para mapear a ação para uma função que `Office.actions.associate` `HIDETASKPANE` chama [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="bd68f-143">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="bd68f-144">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="bd68f-144">The following is an example:</span></span>

    ```javascript
    Office.actions.associate('HIDETASKPANE', function () {
        return Office.addin.hide()
            .then(function () {
                return;
            })
            .catch(function (error) {
                return error.code;
            });
    });
    ```

<span data-ttu-id="bd68f-145">Seguindo as etapas anteriores, o seu add-in alterna a visibilidade do painel de tarefas pressionando **Ctrl+Alt+Up** e **Ctrl+Alt+Down.**</span><span class="sxs-lookup"><span data-stu-id="bd68f-145">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Alt+Up** and **Ctrl+Alt+Down**.</span></span> <span data-ttu-id="bd68f-146">O mesmo comportamento é mostrado no exemplo Excel [atalhos](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) de teclado no Office PnP de GitHub.</span><span class="sxs-lookup"><span data-stu-id="bd68f-146">The same behavior is shown in the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample in the Office Add-ins PnP repo in GitHub.</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="bd68f-147">Detalhes e restrições</span><span class="sxs-lookup"><span data-stu-id="bd68f-147">Details and restrictions</span></span>

### <a name="construct-the-action-objects"></a><span data-ttu-id="bd68f-148">Construir os objetos de ação</span><span class="sxs-lookup"><span data-stu-id="bd68f-148">Construct the action objects</span></span>

<span data-ttu-id="bd68f-149">Use as seguintes diretrizes ao especificar os objetos na matriz do shortcuts.js`actions` on:</span><span class="sxs-lookup"><span data-stu-id="bd68f-149">Use the following guidelines when specifying the objects in the `actions` array of the shortcuts.json:</span></span>

- <span data-ttu-id="bd68f-150">Os nomes das `id` propriedades e `name` são obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="bd68f-150">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="bd68f-151">A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-151">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="bd68f-152">A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação.</span><span class="sxs-lookup"><span data-stu-id="bd68f-152">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="bd68f-153">Deve ser uma combinação dos caracteres A - Z, a - z, 0 - 9 e as marcas de pontuação "-", "_" e "+".</span><span class="sxs-lookup"><span data-stu-id="bd68f-153">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="bd68f-154">A propriedade do `type` é opcional.</span><span class="sxs-lookup"><span data-stu-id="bd68f-154">The `type` property is optional.</span></span> <span data-ttu-id="bd68f-155">Atualmente, apenas `ExecuteFunction` o tipo é suportado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-155">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="bd68f-156">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="bd68f-156">The following is an example:</span></span>

```json
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        },
        {
            "id": "HIDETASKPANE",
            "type": "ExecuteFunction",
            "name": "Hide task pane for add-in"
        }
    ]
```

<span data-ttu-id="bd68f-157">O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="bd68f-157">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="construct-the-shortcut-objects"></a><span data-ttu-id="bd68f-158">Construir os objetos de atalho</span><span class="sxs-lookup"><span data-stu-id="bd68f-158">Construct the shortcut objects</span></span>

<span data-ttu-id="bd68f-159">Use as seguintes diretrizes ao especificar os objetos na matriz do shortcuts.js`shortcuts` on:</span><span class="sxs-lookup"><span data-stu-id="bd68f-159">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="bd68f-160">Os nomes de `action` propriedade , e são `key` `default` obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="bd68f-160">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="bd68f-161">O valor da propriedade é uma cadeia de caracteres e deve corresponder a uma das `action` `id` propriedades no objeto action.</span><span class="sxs-lookup"><span data-stu-id="bd68f-161">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="bd68f-162">A propriedade pode ser qualquer combinação dos caracteres A - Z, a -z, 0 - 9 e as marcas de pontuação `default` "-", "_" e "+".</span><span class="sxs-lookup"><span data-stu-id="bd68f-162">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="bd68f-163">(Por convenção, letras de maiúsculas e baixos não são usadas nessas propriedades.)</span><span class="sxs-lookup"><span data-stu-id="bd68f-163">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="bd68f-164">A propriedade deve conter o nome de pelo menos uma chave `default` modificadora (Alt, Ctrl, Shift) e apenas uma outra chave.</span><span class="sxs-lookup"><span data-stu-id="bd68f-164">The `default` property must contain the name of at least one modifier key (Alt, Ctrl, Shift) and only one other key.</span></span> 
- <span data-ttu-id="bd68f-165">Shift não pode ser usado como a única chave modificadora.</span><span class="sxs-lookup"><span data-stu-id="bd68f-165">Shift cannot be used as the only modifier key.</span></span> <span data-ttu-id="bd68f-166">Combine Shift com Alt ou Ctrl.</span><span class="sxs-lookup"><span data-stu-id="bd68f-166">Combine Shift with either Alt or Ctrl.</span></span>
- <span data-ttu-id="bd68f-167">Para Macs, também há suporte para a chave do modificador de comando.</span><span class="sxs-lookup"><span data-stu-id="bd68f-167">For Macs, we also support the Command modifier key.</span></span>
- <span data-ttu-id="bd68f-168">Para Macs, Alt é mapeado para a tecla Option.</span><span class="sxs-lookup"><span data-stu-id="bd68f-168">For Macs, Alt is mapped to the Option key.</span></span> <span data-ttu-id="bd68f-169">Para Windows, Command é mapeado para a tecla Ctrl.</span><span class="sxs-lookup"><span data-stu-id="bd68f-169">For Windows, Command is mapped to the Ctrl key.</span></span>
- <span data-ttu-id="bd68f-170">Quando dois caracteres são vinculados à mesma chave física em um teclado padrão, eles são sinônimos na propriedade; por exemplo, Alt+a e Alt+A são o mesmo atalho, assim como `default` Ctrl+- e Ctrl+ porque "-" e "_" são a mesma chave \_ física.</span><span class="sxs-lookup"><span data-stu-id="bd68f-170">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, Alt+a and Alt+A are the same shortcut, so are Ctrl+- and Ctrl+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="bd68f-171">O caractere "+" indica que as teclas de cada lado são pressionadas simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="bd68f-171">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="bd68f-172">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="bd68f-172">The following is an example:</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down"
            }
        }
    ]
```

<span data-ttu-id="bd68f-173">O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="bd68f-173">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="bd68f-174">As Dicas de Chave, também conhecidas como atalhos de chave sequencial, como o atalho Excel para escolher uma cor de preenchimento **Alt+H, H**, não são suportadas em Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="bd68f-174">KeyTips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a><span data-ttu-id="bd68f-175">Evitar combinações de teclas em uso por outros complementos</span><span class="sxs-lookup"><span data-stu-id="bd68f-175">Avoid key combinations in use by other add-ins</span></span>

<span data-ttu-id="bd68f-176">Há muitos atalhos de teclado que já estão em uso por Office.</span><span class="sxs-lookup"><span data-stu-id="bd68f-176">There are many keyboard shortcuts that are already in use by Office.</span></span> <span data-ttu-id="bd68f-177">Evite registrar atalhos de teclado para o seu complemento que já estão em uso, no entanto, pode haver algumas instâncias em que é necessário substituir atalhos de teclado existentes ou lidar com conflitos entre vários complementos que registraram o mesmo atalho de teclado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-177">Avoid registering keyboard shortcuts for your add-in that are already in use, however there may be some instances where it is necessary to override existing keyboard shortcuts or handle conflicts between multiple add-ins that have registered the same keyboard shortcut.</span></span>

<span data-ttu-id="bd68f-178">No caso de um conflito, o usuário verá uma caixa de diálogo na primeira vez que tentar usar um atalho de teclado conflitante, observe que o nome da ação exibido nesta caixa de diálogo é a propriedade no objeto action no `name` `shortcuts.json` arquivo.</span><span class="sxs-lookup"><span data-stu-id="bd68f-178">In the case of a conflict, the user will see a dialog box the first time they attempt to use a conflicting keyboard shortcut, note that the action name that is displayed in this dialog is the `name` property in the action object in `shortcuts.json` file.</span></span>

![Ilustração mostrando um modo de conflito com duas ações diferentes para um único atalho](../images/add-in-shortcut-conflict-modal.png)

<span data-ttu-id="bd68f-180">O usuário pode selecionar qual ação o atalho do teclado tomará.</span><span class="sxs-lookup"><span data-stu-id="bd68f-180">The user can select which action the keyboard shortcut will take.</span></span> <span data-ttu-id="bd68f-181">Depois de fazer a seleção, a preferência é salva para usos futuros do mesmo atalho.</span><span class="sxs-lookup"><span data-stu-id="bd68f-181">After making the selection, the preference is saved for future uses of the same shortcut.</span></span> <span data-ttu-id="bd68f-182">As preferências de atalho são salvas por usuário, por plataforma.</span><span class="sxs-lookup"><span data-stu-id="bd68f-182">The shortcut preferences are saved per user, per platform.</span></span> <span data-ttu-id="bd68f-183">Se o usuário desejar alterar suas preferências, poderá invocar o comando Redefinir Office preferências de atalho de **Complementos** da caixa de pesquisa Diga-me. </span><span class="sxs-lookup"><span data-stu-id="bd68f-183">If the user wishes to change their preferences, they can invoke the **Reset Office Add-ins shortcut preferences** command from the **Tell me** search box.</span></span> <span data-ttu-id="bd68f-184">Invocar o comando limpa todas as preferências de atalho do complemento do usuário e o usuário será novamente solicitado com a caixa de diálogo conflito na próxima vez que tentar usar um atalho conflitante:</span><span class="sxs-lookup"><span data-stu-id="bd68f-184">Invoking the command clears all of the user's add-in shortcut preferences and the user will again be prompted with the conflict dialog box the next time they attempt to use a conflicting shortcut:</span></span>

![A caixa de pesquisa Diga-me no Excel mostrando a ação redefinir Office preferências de atalho do add-in](../images/add-in-reset-shortcuts-action.png)

<span data-ttu-id="bd68f-186">Para a melhor experiência do usuário, recomendamos que você minimize os conflitos com Excel com essas boas práticas:</span><span class="sxs-lookup"><span data-stu-id="bd68f-186">For the best user experience, we recommend that you minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="bd68f-187">Use apenas atalhos de teclado com o seguinte padrão: \**Ctrl+Shift+Alt+* x\*\*\*, onde *x* é alguma outra chave.</span><span class="sxs-lookup"><span data-stu-id="bd68f-187">Use only keyboard shortcuts with the following pattern: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="bd68f-188">Se você precisar de mais atalhos de teclado, verifique a lista de [atalhos](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)Excel teclado e evite usar qualquer um deles no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="bd68f-188">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>
- <span data-ttu-id="bd68f-189">Quando o foco do teclado estiver dentro da interface do usuário do complemento, **Ctrl+Spacebar** e **Ctrl+Shift+F10** não funcionarão, pois são atalhos de acessibilidade essenciais.</span><span class="sxs-lookup"><span data-stu-id="bd68f-189">When the keyboard focus is inside the add-in UI, **Ctrl+Spacebar** and **Ctrl+Shift+F10** will not work as these are essential accessibility shortcuts.</span></span>
- <span data-ttu-id="bd68f-190">Em um computador Windows ou Mac, se o comando "Redefinir preferências de atalho de complementos do Office" não estiver disponível no menu de pesquisa, o usuário poderá adicionar manualmente o comando à faixa de opções personalização da faixa de opções por meio do menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="bd68f-190">On a Windows or Mac computer, if the "Reset Office Add-ins shortcut preferences" command is not available on the search menu, the user can manually add the command to the ribbon by customizing the ribbon through the context menu.</span></span>

## <a name="customize-the-keyboard-shortcuts-per-platform"></a><span data-ttu-id="bd68f-191">Personalizar os atalhos de teclado por plataforma</span><span class="sxs-lookup"><span data-stu-id="bd68f-191">Customize the keyboard shortcuts per platform</span></span>

<span data-ttu-id="bd68f-192">É possível personalizar atalhos para serem específicos da plataforma.</span><span class="sxs-lookup"><span data-stu-id="bd68f-192">It's possible to customize shortcuts to be platform-specific.</span></span> <span data-ttu-id="bd68f-193">Veja a seguir um exemplo do objeto que personaliza os atalhos para cada uma das seguintes `shortcuts` plataformas: `windows` , `mac` , `web` .</span><span class="sxs-lookup"><span data-stu-id="bd68f-193">The following is an example of the `shortcuts` object that customizes the shortcuts for each of the following platforms: `windows`, `mac`, `web`.</span></span> <span data-ttu-id="bd68f-194">Observe que você ainda deve ter uma tecla `default` de atalho para cada atalho.</span><span class="sxs-lookup"><span data-stu-id="bd68f-194">Note that you must still have a `default` shortcut key for each shortcut.</span></span>

<span data-ttu-id="bd68f-195">No exemplo a seguir, a chave é a chave `default` de fallback para qualquer plataforma que não seja especificada.</span><span class="sxs-lookup"><span data-stu-id="bd68f-195">In the following example, the `default` key is the fallback key for any platform that is not specified.</span></span> <span data-ttu-id="bd68f-196">A única plataforma não especificada é Windows, portanto, a `default` chave só se aplicará a Windows.</span><span class="sxs-lookup"><span data-stu-id="bd68f-196">The only platform not specified is Windows, so the `default` key will only apply to Windows.</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "Ctrl+Alt+Up",
                "mac": "Command+Shift+Up",
                "web": "Ctrl+Alt+1",
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "Ctrl+Alt+Down",
                "mac": "Command+Shift+Down",
                "web": "Ctrl+Alt+2"
            }
        }
    ]
```

## <a name="localize-the-keyboard-shortcuts-json"></a><span data-ttu-id="bd68f-197">Localize os atalhos de teclado JSON</span><span class="sxs-lookup"><span data-stu-id="bd68f-197">Localize the keyboard shortcuts JSON</span></span>

<span data-ttu-id="bd68f-198">Se o seu add-in dá suporte a várias localidades, você precisará localizar a `name` propriedade dos objetos de ação.</span><span class="sxs-lookup"><span data-stu-id="bd68f-198">If your add-in supports multiple locales, you'll need to localize the `name` property of the action objects.</span></span> <span data-ttu-id="bd68f-199">Além disso, se qualquer uma das localidades com suporte para o complemento tiver alfabetos ou sistemas de escrita diferentes e, portanto, teclados diferentes, talvez seja necessário localizar os atalhos também.</span><span class="sxs-lookup"><span data-stu-id="bd68f-199">Also, if any of the locales that the add-in supports have alphabets or different writing systems, and hence different keyboards, you may need to localize the shortcuts also.</span></span> <span data-ttu-id="bd68f-200">Para obter informações sobre como localizar os atalhos de teclado JSON, consulte [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span><span class="sxs-lookup"><span data-stu-id="bd68f-200">For information about how to localize the keyboard shortcuts JSON, see [Localize extended overrides](../develop/localization.md#localize-extended-overrides).</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="bd68f-201">Atalhos do navegador que não podem ser substituídos</span><span class="sxs-lookup"><span data-stu-id="bd68f-201">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="bd68f-202">Ao usar atalhos de teclado personalizados na Web, alguns atalhos de teclado usados pelo navegador não podem ser substituídos por complementos. Esta lista é um trabalho em andamento.</span><span class="sxs-lookup"><span data-stu-id="bd68f-202">When using custom keyboard shortcuts on the web, some keyboard shortcuts that are used by the browser cannot be overridden by add-ins. This list is a work in progress.</span></span> <span data-ttu-id="bd68f-203">Se você descobrir outras combinações que não podem ser substituidas, nos avise usando a ferramenta de comentários na parte inferior desta página.</span><span class="sxs-lookup"><span data-stu-id="bd68f-203">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="bd68f-204">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="bd68f-204">Ctrl+N</span></span>
- <span data-ttu-id="bd68f-205">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="bd68f-205">Ctrl+Shift+N</span></span>
- <span data-ttu-id="bd68f-206">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="bd68f-206">Ctrl+T</span></span>
- <span data-ttu-id="bd68f-207">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="bd68f-207">Ctrl+Shift+T</span></span>
- <span data-ttu-id="bd68f-208">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="bd68f-208">Ctrl+W</span></span>
- <span data-ttu-id="bd68f-209">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="bd68f-209">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="bd68f-210">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="bd68f-210">Next Steps</span></span>

- <span data-ttu-id="bd68f-211">Consulte o [Excel de exemplo de atalhos](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) de teclado.</span><span class="sxs-lookup"><span data-stu-id="bd68f-211">See the [Excel keyboard shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) sample add-in.</span></span>
- <span data-ttu-id="bd68f-212">Obter uma visão geral de como trabalhar com substituições estendidas em [Trabalho com substituições estendidas do manifesto](../develop/extended-overrides.md).</span><span class="sxs-lookup"><span data-stu-id="bd68f-212">Get an overview of working with extended overrides in [Work with extended overrides of the manifest](../develop/extended-overrides.md).</span></span>
