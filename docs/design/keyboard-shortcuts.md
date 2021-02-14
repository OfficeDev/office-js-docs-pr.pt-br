---
title: Atalhos de teclado personalizados em Complementos do Office
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao seu Complemento do Office.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: 7374c23bacad031024edc7b84da951a8e91fc175
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237683"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a><span data-ttu-id="efd72-103">Adicionar atalhos de teclado personalizados aos seus Complementos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="efd72-103">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>

<span data-ttu-id="efd72-104">Os atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do seu complemento trabalhem com mais eficiência e melhoram a acessibilidade do complemento para usuários com deficiências, fornecendo uma alternativa ao mouse.</span><span class="sxs-lookup"><span data-stu-id="efd72-104">Keyboard shortcuts, also known as key combinations, enable your add-in's users to work more efficiently and they improve the add-in's accessibility for users with disabilities by providing an alternative to the mouse.</span></span>

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> <span data-ttu-id="efd72-105">Para começar com uma versão de trabalho de um complemento com atalhos de teclado já habilitados, clone e execute o exemplo de atalhos de teclado [do Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="efd72-105">To start with a working version of an add-in with keyboard shortcuts already enabled, clone and run the sample [Excel Keyboard Shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span> <span data-ttu-id="efd72-106">Quando você estiver pronto para adicionar atalhos de teclado ao seu próprio complemento, continue com este artigo.</span><span class="sxs-lookup"><span data-stu-id="efd72-106">When you are ready to add keyboard shortcuts to your own add-in, continue with this article.</span></span>

<span data-ttu-id="efd72-107">Há três etapas para adicionar atalhos de teclado a um complemento:</span><span class="sxs-lookup"><span data-stu-id="efd72-107">There are three steps to add keyboard shortcuts to an add-in:</span></span>

1. <span data-ttu-id="efd72-108">[Configure o manifesto do complemento.](#configure-the-manifest)</span><span class="sxs-lookup"><span data-stu-id="efd72-108">[Configure the add-in's manifest](#configure-the-manifest).</span></span>
1. <span data-ttu-id="efd72-109">[Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e seus atalhos de teclado.</span><span class="sxs-lookup"><span data-stu-id="efd72-109">[Create or edit the shortcuts JSON file](#create-or-edit-the-shortcuts-json-file) to define actions and their keyboard shortcuts.</span></span>
1. <span data-ttu-id="efd72-110">[Adicione uma ou mais chamadas de tempo de](#create-a-mapping-of-actions-to-their-functions) execução da API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear uma função para cada ação.</span><span class="sxs-lookup"><span data-stu-id="efd72-110">[Add one or more runtime calls](#create-a-mapping-of-actions-to-their-functions) of the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map a function to each action.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="efd72-111">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="efd72-111">Configure the manifest</span></span>

<span data-ttu-id="efd72-112">Há duas pequenas alterações no manifesto a fazer.</span><span class="sxs-lookup"><span data-stu-id="efd72-112">There are two small changes to the manifest to make.</span></span> <span data-ttu-id="efd72-113">Uma é habilitar o add-in para usar um tempo de execução compartilhado e a outra é apontar para um arquivo formatado em JSON onde você definiu os atalhos de teclado.</span><span class="sxs-lookup"><span data-stu-id="efd72-113">One is to enable the add-in to use a shared runtime and the other is to point to a JSON-formatted file where you defined the keyboard shortcuts.</span></span>

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="efd72-114">Configurar o complemento para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="efd72-114">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="efd72-115">A adição de atalhos de teclado personalizados exige que o seu complemento use o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="efd72-115">Adding custom keyboard shortcuts requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="efd72-116">Para obter mais informações, [configure um complemento para usar um tempo de execução compartilhado.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="efd72-116">For more information, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

### <a name="link-the-mapping-file-to-the-manifest"></a><span data-ttu-id="efd72-117">Vincular o arquivo de mapeamento ao manifesto</span><span class="sxs-lookup"><span data-stu-id="efd72-117">Link the mapping file to the manifest</span></span>

<span data-ttu-id="efd72-118">Imediatamente *abaixo* (não dentro) `<VersionOverrides>` do elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="efd72-118">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="efd72-119">De definir o atributo para a URL completa de um arquivo JSON em seu projeto `Url` que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="efd72-119">Set the `Url` attribute to the full URL of a JSON file in your project that you will create in a later step.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a><span data-ttu-id="efd72-120">Criar ou editar o arquivo JSON de atalhos</span><span class="sxs-lookup"><span data-stu-id="efd72-120">Create or edit the shortcuts JSON file</span></span>

<span data-ttu-id="efd72-121">Crie um arquivo JSON em seu projeto.</span><span class="sxs-lookup"><span data-stu-id="efd72-121">Create a JSON file in your project.</span></span> <span data-ttu-id="efd72-122">Certifique-se de que o caminho do arquivo corresponde ao local especificado para `Url` o atributo do elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="efd72-122">Be sure the path of the file matches the location you specified for the `Url` attribute of the [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="efd72-123">Esse arquivo descreverá os atalhos de teclado e as ações que eles chamarão.</span><span class="sxs-lookup"><span data-stu-id="efd72-123">This file will describe your keyboard shortcuts, and the actions that they will invoke.</span></span>

1. <span data-ttu-id="efd72-124">Dentro do arquivo JSON, há duas matrizes.</span><span class="sxs-lookup"><span data-stu-id="efd72-124">Inside the JSON file, there are two arrays.</span></span> <span data-ttu-id="efd72-125">A matriz de ações conterá objetos que definem as ações a serem invocadas e a matriz de atalhos conterá objetos que mapeiam combinações de teclas em ações.</span><span class="sxs-lookup"><span data-stu-id="efd72-125">The actions array will contain objects that define the actions to be invoked and the shortcuts array will contain objects that map key combinations onto actions.</span></span> <span data-ttu-id="efd72-126">Veja um exemplo:</span><span class="sxs-lookup"><span data-stu-id="efd72-126">Here is an example:</span></span>

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
                    "default": "CTRL+SHIFT+UP"
                }
            },
            {
                "action": "HIDETASKPANE",
                "key": {
                    "default": "CTRL+SHIFT+DOWN"
                }
            }
        ]
    }
    ```

    <span data-ttu-id="efd72-127">Para obter mais informações sobre os objetos JSON, consulte Construindo os objetos [de ação](#constructing-the-action-objects) e construindo os objetos [de atalho.](#constructing-the-shortcut-objects)</span><span class="sxs-lookup"><span data-stu-id="efd72-127">For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects).</span></span> <span data-ttu-id="efd72-128">O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="efd72-128">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

    > [!NOTE]
    > <span data-ttu-id="efd72-129">Você pode usar "CONTROL" no lugar de "CTRL" neste artigo.</span><span class="sxs-lookup"><span data-stu-id="efd72-129">You can use "CONTROL" in place of "CTRL" throughout this article.</span></span>

    <span data-ttu-id="efd72-130">Em uma etapa posterior, as ações serão mapeadas para as funções que você escrever.</span><span class="sxs-lookup"><span data-stu-id="efd72-130">In a later step, the actions will themselves be mapped to functions that you write.</span></span> <span data-ttu-id="efd72-131">Neste exemplo, você mapeará SHOWTASKPANE para uma função que chama o método e HIDETASKPANE para uma função que chama `Office.addin.showAsTaskpane` o `Office.addin.hide` método.</span><span class="sxs-lookup"><span data-stu-id="efd72-131">In this example, you will later map SHOWTASKPANE to a function that calls the `Office.addin.showAsTaskpane` method and HIDETASKPANE to a function that calls the `Office.addin.hide` method.</span></span>

## <a name="create-a-mapping-of-actions-to-their-functions"></a><span data-ttu-id="efd72-132">Criar um mapeamento de ações para suas funções</span><span class="sxs-lookup"><span data-stu-id="efd72-132">Create a mapping of actions to their functions</span></span>

1. <span data-ttu-id="efd72-133">No projeto, abra o arquivo JavaScript carregado pela página HTML no `<FunctionFile>` elemento.</span><span class="sxs-lookup"><span data-stu-id="efd72-133">In your project, open the JavaScript file loaded by your HTML page in the `<FunctionFile>` element.</span></span>
1. <span data-ttu-id="efd72-134">No arquivo JavaScript, use a API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear cada ação especificada no arquivo JSON para uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="efd72-134">In the JavaScript file, use the [Office.actions.associate](/javascript/api/office/office.actions#associate) API to map each action that you specified in the JSON file to a JavaScript function.</span></span> <span data-ttu-id="efd72-135">Adicione o JavaScript a seguir ao arquivo.</span><span class="sxs-lookup"><span data-stu-id="efd72-135">Add the following JavaScript to the file.</span></span> <span data-ttu-id="efd72-136">Observe o seguinte sobre o código:</span><span class="sxs-lookup"><span data-stu-id="efd72-136">Note the following about the code:</span></span>

    - <span data-ttu-id="efd72-137">O primeiro parâmetro é uma das ações do arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="efd72-137">The first parameter is one of the actions from the JSON file.</span></span>
    - <span data-ttu-id="efd72-138">O segundo parâmetro é a função que é executado quando um usuário pressiona a combinação de teclas mapeada para a ação no arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="efd72-138">The second parameter is the function that runs when a user presses the key combination that is mapped to the action in the JSON file.</span></span>

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. <span data-ttu-id="efd72-139">Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="efd72-139">To continue the example, use `'SHOWTASKPANE'` as the first parameter.</span></span>
1. <span data-ttu-id="efd72-140">Para o corpo da função, use o [método Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) para abrir o painel de tarefas do complemento.</span><span class="sxs-lookup"><span data-stu-id="efd72-140">For the body of the function, use the [Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) method to open the add-in's task pane.</span></span> <span data-ttu-id="efd72-141">Quando terminar, o código deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="efd72-141">When you are done, the code should look like the following:</span></span>

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

1. <span data-ttu-id="efd72-142">Adicione uma segunda chamada de função para mapear a ação para uma função que chama `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span><span class="sxs-lookup"><span data-stu-id="efd72-142">Add a second call of `Office.actions.associate` function to map the `HIDETASKPANE` action to a function that calls [Office.addin.hide](/javascript/api/office/office.addin#hide--).</span></span> <span data-ttu-id="efd72-143">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="efd72-143">The following is an example:</span></span>

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

<span data-ttu-id="efd72-144">Seguir as etapas anteriores permite que o seu complemento alterne a visibilidade do painel de tarefas pressionando **Ctrl+Shift+Tecla** de seta para cima e Tecla de seta **Ctrl+Shift+Down.**</span><span class="sxs-lookup"><span data-stu-id="efd72-144">Following the previous steps lets your add-in toggle the visibility of the task pane by pressing **Ctrl+Shift+Up arrow key** and **Ctrl+Shift+Down arrow key**.</span></span> <span data-ttu-id="efd72-145">Esse é o mesmo comportamento mostrado no exemplo de atalhos de teclado [do Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)</span><span class="sxs-lookup"><span data-stu-id="efd72-145">This is the same behavior as shown in the [sample excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>

## <a name="details-and-restrictions"></a><span data-ttu-id="efd72-146">Detalhes e restrições</span><span class="sxs-lookup"><span data-stu-id="efd72-146">Details and restrictions</span></span>

### <a name="constructing-the-action-objects"></a><span data-ttu-id="efd72-147">Construir os objetos de ação</span><span class="sxs-lookup"><span data-stu-id="efd72-147">Constructing the action objects</span></span>

<span data-ttu-id="efd72-148">Use as diretrizes a seguir ao especificar os objetos na `action` matriz do shortcuts.jsem:</span><span class="sxs-lookup"><span data-stu-id="efd72-148">Use the following guidelines when specifying the objects in the `action` array of the shortcuts.json:</span></span>

- <span data-ttu-id="efd72-149">Os nomes das `id` propriedades e `name` são obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="efd72-149">The property names `id` and `name` are mandatory.</span></span>
- <span data-ttu-id="efd72-150">A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.</span><span class="sxs-lookup"><span data-stu-id="efd72-150">The `id` property is used to uniquely identify the action to invoke using a keyboard shortcut.</span></span>
- <span data-ttu-id="efd72-151">A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação.</span><span class="sxs-lookup"><span data-stu-id="efd72-151">The `name` property must be a user friendly string describing the action.</span></span> <span data-ttu-id="efd72-152">Deve ser uma combinação dos caracteres A - Z, a - z, 0 - 9 e as marcas de pontuação "-", "_" e "+".</span><span class="sxs-lookup"><span data-stu-id="efd72-152">It must be a combination of the characters A - Z, a - z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span>
- <span data-ttu-id="efd72-153">A propriedade do `type` é opcional.</span><span class="sxs-lookup"><span data-stu-id="efd72-153">The `type` property is optional.</span></span> <span data-ttu-id="efd72-154">Atualmente, só `ExecuteFunction` há suporte para o tipo.</span><span class="sxs-lookup"><span data-stu-id="efd72-154">Currently only `ExecuteFunction` type is supported.</span></span>

<span data-ttu-id="efd72-155">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="efd72-155">The following is an example:</span></span>

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

<span data-ttu-id="efd72-156">O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="efd72-156">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

### <a name="constructing-the-shortcut-objects"></a><span data-ttu-id="efd72-157">Construir os objetos de atalho</span><span class="sxs-lookup"><span data-stu-id="efd72-157">Constructing the shortcut objects</span></span>

<span data-ttu-id="efd72-158">Use as diretrizes a seguir ao especificar os objetos na `shortcuts` matriz do shortcuts.jsem:</span><span class="sxs-lookup"><span data-stu-id="efd72-158">Use the following guidelines when specifying the objects in the `shortcuts` array of the shortcuts.json:</span></span>

- <span data-ttu-id="efd72-159">Os nomes de `action` propriedade , e são `key` `default` obrigatórios.</span><span class="sxs-lookup"><span data-stu-id="efd72-159">The property names `action`, `key`, and `default` are required.</span></span>
- <span data-ttu-id="efd72-160">O valor da propriedade `action` é uma cadeia de caracteres e deve corresponder a uma das propriedades no objeto `id` action.</span><span class="sxs-lookup"><span data-stu-id="efd72-160">The value of the `action` property is a string and must match one of the `id` properties in the action object.</span></span>
- <span data-ttu-id="efd72-161">A propriedade pode ser qualquer combinação dos caracteres A - Z, a -z, 0 - 9 e as marcas de `default` pontuação "-", "_" e "+".</span><span class="sxs-lookup"><span data-stu-id="efd72-161">The `default` property can be any combination of the characters A - Z, a -z, 0 - 9, and the punctuation marks "-", "_", and "+".</span></span> <span data-ttu-id="efd72-162">(Por convenção, letras maiúsculas e baixas não são usadas nessas propriedades.)</span><span class="sxs-lookup"><span data-stu-id="efd72-162">(By convention, lower case letters are not used in these properties.)</span></span>
- <span data-ttu-id="efd72-163">A propriedade deve conter o nome de pelo menos uma tecla modificadora `default` (ALT, CTRL, SHIFT) e apenas uma outra tecla.</span><span class="sxs-lookup"><span data-stu-id="efd72-163">The `default` property must contain the name of at least one modifier key (ALT, CTRL, SHIFT) and only one other key.</span></span>
- <span data-ttu-id="efd72-164">Para Macs, também damos suporte à tecla modificadora COMMAND.</span><span class="sxs-lookup"><span data-stu-id="efd72-164">For Macs, we also support the COMMAND modifier key.</span></span>
- <span data-ttu-id="efd72-165">Para Macs, ALT é mapeada para a tecla OPTION.</span><span class="sxs-lookup"><span data-stu-id="efd72-165">For Macs, ALT is mapped to the OPTION key.</span></span> <span data-ttu-id="efd72-166">Para o Windows, COMMAND é mapeado para a tecla CTRL.</span><span class="sxs-lookup"><span data-stu-id="efd72-166">For Windows, COMMAND is mapped to the CTRL key.</span></span>
- <span data-ttu-id="efd72-167">Quando dois caracteres são vinculados à mesma tecla física em um teclado padrão, eles são sinônimos na propriedade; por exemplo, ALT+a e ALT+A são o mesmo atalho, assim como `default` CTRL+- e CTRL+ porque "-" e "_" são a mesma tecla \_ física.</span><span class="sxs-lookup"><span data-stu-id="efd72-167">When two characters are linked to the same physical key in a standard keyboard, then they are synonyms in the `default` property; for example, ALT+a and ALT+A are the same shortcut, so are CTRL+- and CTRL+\_ because "-" and "_" are the same physical key.</span></span>
- <span data-ttu-id="efd72-168">O caractere "+" indica que as teclas de cada lado são pressionadas simultaneamente.</span><span class="sxs-lookup"><span data-stu-id="efd72-168">The "+" character indicates that the keys on either side of it are pressed simultaneously.</span></span>

<span data-ttu-id="efd72-169">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="efd72-169">The following is an example:</span></span>

```json
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+UP"
            }
        },
        {
            "action": "HIDETASKPANE",
            "key": {
                "default": "CTRL+SHIFT+DOWN"
            }
        }
    ]
```

<span data-ttu-id="efd72-170">O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="efd72-170">The complete schema for the shortcuts JSON is at [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!NOTE]
> <span data-ttu-id="efd72-171">Dicas de tecla, também conhecidas como atalhos de teclas sequenciais, como o atalho do Excel para escolher uma cor de preenchimento **Alt+H, H**, não são suportadas nos Complementos do Office.</span><span class="sxs-lookup"><span data-stu-id="efd72-171">Keytips, also known as sequential key shortcuts, such as the Excel shortcut to choose a fill color **Alt+H, H**, are not supported in Office Add-ins.</span></span>

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a><span data-ttu-id="efd72-172">Usando atalhos quando o foco está no painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="efd72-172">Using shortcuts when the focus is in the task pane</span></span>

<span data-ttu-id="efd72-173">Atualmente, os atalhos de teclado para um Complemento do Office só podem ser invocados quando o foco do usuário está na planilha.</span><span class="sxs-lookup"><span data-stu-id="efd72-173">Currently, the keyboard shortcuts for an Office Add-in can only be invoked when the user's focus is in the worksheet.</span></span> <span data-ttu-id="efd72-174">Quando o foco do usuário está dentro da interface do usuário do Office (como o painel de tarefas), nenhum dos atalhos do complemento é ignorado.</span><span class="sxs-lookup"><span data-stu-id="efd72-174">When the user's focus is inside the Office UI (such as the task pane), none of the add-in's shortcuts are ignored.</span></span> <span data-ttu-id="efd72-175">Como alternativa, o complemento pode definir manipuladores de teclado que podem invocar determinadas ações quando o foco do usuário está dentro da interface do usuário do complemento.</span><span class="sxs-lookup"><span data-stu-id="efd72-175">As a workaround, the add-in can define keyboard handlers that can invoke certain actions when the user's focus is inside of the add-in UI.</span></span>

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a><span data-ttu-id="efd72-176">Usando combinações de teclas que já são usadas pelo Office ou outro complemento</span><span class="sxs-lookup"><span data-stu-id="efd72-176">Using key combinations that are already used by Office or another add-in</span></span>

<span data-ttu-id="efd72-177">Durante o período de visualização, não há sistema para determinar o que acontece quando um usuário pressiona uma combinação de teclas registrada por um complemento e também pelo Office ou por outro complemento.</span><span class="sxs-lookup"><span data-stu-id="efd72-177">During the preview period, there is no system for determining what happens when a user presses a key combination that is registered by an add-in and also by Office or by another add-in.</span></span> <span data-ttu-id="efd72-178">O comportamento é indefinido.</span><span class="sxs-lookup"><span data-stu-id="efd72-178">Behavior is undefined.</span></span>

<span data-ttu-id="efd72-179">Atualmente, não há uma solução alternativa quando dois ou mais complementos registraram o mesmo atalho de teclado, mas você pode minimizar conflitos com o Excel com estas práticas recomendadas:</span><span class="sxs-lookup"><span data-stu-id="efd72-179">Currently, there is no workaround when two or more add-ins have registered the same keyboard shortcut, but you can minimize conflicts with Excel with these good practices:</span></span>

- <span data-ttu-id="efd72-180">Use somente atalhos de teclado com o seguinte padrão no seu complemento: \**Ctrl+Shift+Alt+* x\*\*\*, onde *x* é alguma outra tecla.</span><span class="sxs-lookup"><span data-stu-id="efd72-180">Use only keyboard shortcuts with the following pattern in your add-in: \**Ctrl+Shift+Alt+* x\*\*\*, where *x* is some other key.</span></span>
- <span data-ttu-id="efd72-181">Se precisar de mais atalhos de teclado, verifique a lista de [atalhos](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)de teclado do Excel e evite usar qualquer um deles no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="efd72-181">If you need more keyboard shortcuts, check the [list of Excel keyboard shortcuts](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f), and avoid using any of them in your add-in.</span></span>

## <a name="browser-shortcuts-that-cannot-be-overridden"></a><span data-ttu-id="efd72-182">Atalhos do navegador que não podem ser substituídos</span><span class="sxs-lookup"><span data-stu-id="efd72-182">Browser shortcuts that cannot be overridden</span></span>

<span data-ttu-id="efd72-183">Você não pode usar nenhuma das combinações de teclado a seguir.</span><span class="sxs-lookup"><span data-stu-id="efd72-183">You cannot use any of the following keyboard combinations.</span></span> <span data-ttu-id="efd72-184">Eles são usados por navegadores e não podem ser substituídos.</span><span class="sxs-lookup"><span data-stu-id="efd72-184">They are used by browsers and cannot be overridden.</span></span> <span data-ttu-id="efd72-185">Esta lista é um trabalho em andamento.</span><span class="sxs-lookup"><span data-stu-id="efd72-185">This list is a work in progress.</span></span> <span data-ttu-id="efd72-186">Se você descobrir outras combinações que não podem ser substituídos, diga-nos usando a ferramenta de comentários na parte inferior desta página.</span><span class="sxs-lookup"><span data-stu-id="efd72-186">If you discover other combinations that cannot be overridden, please let us know by using the feedback tool at the bottom of this page.</span></span>

- <span data-ttu-id="efd72-187">Ctrl+N</span><span class="sxs-lookup"><span data-stu-id="efd72-187">Ctrl+N</span></span>
- <span data-ttu-id="efd72-188">Ctrl+Shift+N</span><span class="sxs-lookup"><span data-stu-id="efd72-188">Ctrl+Shift+N</span></span>
- <span data-ttu-id="efd72-189">Ctrl+T</span><span class="sxs-lookup"><span data-stu-id="efd72-189">Ctrl+T</span></span>
- <span data-ttu-id="efd72-190">Ctrl+Shift+T</span><span class="sxs-lookup"><span data-stu-id="efd72-190">Ctrl+Shift+T</span></span>
- <span data-ttu-id="efd72-191">Ctrl+W</span><span class="sxs-lookup"><span data-stu-id="efd72-191">Ctrl+W</span></span>
- <span data-ttu-id="efd72-192">Ctrl+PgUp/PgDn</span><span class="sxs-lookup"><span data-stu-id="efd72-192">Ctrl+PgUp/PgDn</span></span>

## <a name="next-steps"></a><span data-ttu-id="efd72-193">Próximas Etapas</span><span class="sxs-lookup"><span data-stu-id="efd72-193">Next Steps</span></span>

- <span data-ttu-id="efd72-194">Veja o exemplo de complemento [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span><span class="sxs-lookup"><span data-stu-id="efd72-194">See the sample add-in [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).</span></span>
