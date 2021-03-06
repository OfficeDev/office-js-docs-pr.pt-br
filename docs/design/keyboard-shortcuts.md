---
title: Atalhos de teclado personalizados em Complementos do Office
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao seu Complemento do Office.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: c767c6d5bc23f0a44422452839cd8bdf87bd8715
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505196"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Adicionar atalhos de teclado personalizados aos seus Complementos do Office (visualização)

Atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do seu complemento funcionem com mais eficiência e melhoram a acessibilidade do complemento para usuários com deficiências, fornecendo uma alternativa ao mouse.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Para começar com uma versão de trabalho de um complemento com atalhos de teclado já habilitados, clone e execute o exemplo atalhos de teclado [do Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Quando você estiver pronto para adicionar atalhos de teclado ao seu próprio complemento, continue com este artigo.

Há três etapas para adicionar atalhos de teclado a um complemento:

1. [Configure o manifesto do complemento](#configure-the-manifest).
1. [Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e atalhos de teclado.
1. [Adicione uma ou mais chamadas de](#create-a-mapping-of-actions-to-their-functions) tempo de execução da API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear uma função para cada ação.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Há duas pequenas alterações no manifesto a fazer. Um deles é habilitar o add-in para usar um tempo de execução compartilhado e o outro é apontar para um arquivo formatado JSON onde você definiu os atalhos do teclado.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o add-in para usar um tempo de execução compartilhado

A adição de atalhos personalizados de teclado exige que o seu complemento use o tempo de execução compartilhado. Para obter mais informações, [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Vincular o arquivo de mapeamento ao manifesto

Imediatamente *abaixo* (não dentro) `<VersionOverrides>` do elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) De definir o atributo como a URL completa de um arquivo JSON em `Url` seu projeto que você criará em uma etapa posterior.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Criar ou editar o arquivo JSON de atalhos

Crie um arquivo JSON em seu projeto. Certifique-se de que o caminho do arquivo corresponde ao local especificado para o `Url` atributo do [elemento ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Este arquivo descreverá seus atalhos de teclado e as ações que eles invocarão.

1. Dentro do arquivo JSON, há duas matrizes. A matriz de ações conterá objetos que definem as ações a serem invocadas e a matriz de atalhos conterá objetos que mapeiam combinações de teclas em ações. Veja um exemplo:

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

    Para obter mais informações sobre os objetos JSON, consulte [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects). O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Você pode usar "CONTROL" no lugar de "CTRL" ao longo deste artigo.

    Em uma etapa posterior, as ações serão mapeadas para as funções que você escrever. Neste exemplo, mais tarde você mapeará SHOWTASKPANE para uma função que chama o método e `Office.addin.showAsTaskpane` HIDETASKPANE para uma função que chama o `Office.addin.hide` método.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Criar um mapeamento de ações para suas funções

1. Em seu projeto, abra o arquivo JavaScript carregado pela sua página HTML no `<FunctionFile>` elemento.
1. No arquivo JavaScript, use a API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear cada ação especificada no arquivo JSON para uma função JavaScript. Adicione o JavaScript a seguir ao arquivo. Observe o seguinte sobre o código:

    - O primeiro parâmetro é uma das ações do arquivo JSON.
    - O segundo parâmetro é a função que é executado quando um usuário pressiona a combinação de teclas mapeada para a ação no arquivo JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.
1. Para o corpo da função, use o [método Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) para abrir o painel de tarefas do complemento. Quando terminar, o código deverá ter a seguinte aparência:

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

1. Adicione uma segunda chamada de `Office.actions.associate` função para mapear a ação para uma função que chama `HIDETASKPANE` [Office.addin.hide](/javascript/api/office/office.addin#hide--). Veja um exemplo a seguir:

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

Seguindo as etapas anteriores, o seu add-in alterna a visibilidade do painel de tarefas pressionando a tecla de seta **Ctrl+Shift+Up** e a tecla de seta **Ctrl+Shift+Down.** Esse é o mesmo comportamento mostrado no exemplo [do excel keyboard shortcuts add-in](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).

## <a name="details-and-restrictions"></a>Detalhes e restrições

### <a name="constructing-the-action-objects"></a>Construir os objetos de ação

Use as seguintes diretrizes ao especificar os objetos na matriz do shortcuts.js`action` on:

- Os nomes das `id` propriedades e `name` são obrigatórios.
- A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.
- A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação. Deve ser uma combinação dos caracteres A - Z, a - z, 0 - 9 e as marcas de pontuação "-", "_" e "+".
- A propriedade do `type` é opcional. Atualmente, apenas `ExecuteFunction` o tipo é suportado.

Veja um exemplo a seguir:

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

O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="constructing-the-shortcut-objects"></a>Construir os objetos de atalho

Use as seguintes diretrizes ao especificar os objetos na matriz do shortcuts.js`shortcuts` on:

- Os nomes de `action` propriedade , e são `key` `default` obrigatórios.
- O valor da propriedade é uma cadeia de caracteres e deve corresponder a uma das `action` `id` propriedades no objeto action.
- A propriedade pode ser qualquer combinação dos caracteres A - Z, a -z, 0 - 9 e as marcas de pontuação `default` "-", "_" e "+". (Por convenção, letras de maiúsculas e baixos não são usadas nessas propriedades.)
- A propriedade deve conter o nome de pelo menos uma chave `default` modificadora (ALT, CTRL, SHIFT) e apenas uma outra chave.
- Para Macs, também suportamos a chave modificadora COMMAND.
- Para Macs, ALT é mapeado para a tecla OPTION. Para o Windows, COMMAND é mapeado para a tecla CTRL.
- Quando dois caracteres são vinculados à mesma chave física em um teclado padrão, eles são sinônimos na propriedade; por exemplo, ALT+a e ALT+A são o mesmo atalho, assim como `default` CTRL+- e CTRL+ porque "-" e "_" são a mesma chave \_ física.
- O caractere "+" indica que as teclas de cada lado são pressionadas simultaneamente.

Veja um exemplo a seguir:

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

O esquema completo para os atalhos JSON estáextended-manifest.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Dicas de chave, também conhecidas como atalhos de chave sequencial, como o atalho do Excel para escolher uma cor de preenchimento **Alt+H, H**, não são suportadas em Complementos do Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Usando atalhos quando o foco está no painel de tarefas

Atualmente, os atalhos de teclado para um Add-in do Office só podem ser invocados quando o foco do usuário está na planilha. Quando o foco do usuário está dentro da interface do usuário do Office (como o painel de tarefas), nenhum dos atalhos do complemento é ignorado. Como solução alternativa, o complemento pode definir manipuladores de teclado que podem invocar determinadas ações quando o foco do usuário está dentro da interface do usuário do complemento.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Usando combinações de teclas que já são usadas pelo Office ou outro complemento

Durante o período de visualização, não há nenhum sistema para determinar o que acontece quando um usuário pressiona uma combinação de teclas registrada por um complemento e também pelo Office ou por outro complemento. O comportamento é indefinido.

Atualmente, não há solução alternativa quando dois ou mais complementos registraram o mesmo atalho de teclado, mas você pode minimizar conflitos com o Excel com essas boas práticas:

- Use apenas atalhos de teclado com o seguinte padrão no seu complemento: **Ctrl+Shift+Alt+* x***, onde *x* é outra chave.
- Se você precisar de mais atalhos de teclado, verifique a lista de [atalhos](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)de teclado do Excel e evite usar qualquer um deles no seu complemento.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Atalhos do navegador que não podem ser substituídos

Não é possível usar nenhuma das seguintes combinações de teclado. Eles são usados por navegadores e não podem ser substituídos. Esta lista é um trabalho em andamento. Se você descobrir outras combinações que não podem ser substituidas, nos avise usando a ferramenta de comentários na parte inferior desta página.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="localize-the-keyboard-shortcuts-json"></a>Localize os atalhos de teclado JSON

Se o seu add-in dá suporte a várias localidades, você precisará localizar a `name` propriedade dos objetos de ação. Além disso, se qualquer uma das localidades com suporte para o complemento tiver alfabetos ou sistemas de escrita diferentes e, portanto, teclados diferentes, talvez seja necessário localizar os atalhos também. Para obter informações sobre como localizar os atalhos de teclado JSON, consulte [Localize extended overrides](../develop/localization.md#localize-extended-overrides).

## <a name="next-steps"></a>Próximas etapas

- Consulte o exemplo de complemento [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
- Obter uma visão geral de como trabalhar com substituições estendidas em [Trabalho com substituições estendidas do manifesto](../develop/extended-overrides.md).
