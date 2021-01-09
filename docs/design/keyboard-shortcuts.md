---
title: Atalhos de teclado personalizados em Complementos do Office
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao seu Complemento do Office.
ms.date: 12/17/2020
localization_priority: Normal
ms.openlocfilehash: dc99674b92ebb415b1d49fb28821d8c2e34c8077
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789146"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Adicionar atalhos de teclado personalizados aos seus Complementos do Office (visualização)

Os atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do seu complemento trabalhem com mais eficiência e melhoram a acessibilidade do complemento para usuários com deficiências, fornecendo uma alternativa ao mouse.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Para começar com uma versão de trabalho de um complemento com atalhos de teclado já habilitados, clone e execute o exemplo de atalhos de teclado [do Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts) Quando você estiver pronto para adicionar atalhos de teclado ao seu próprio complemento, continue com este artigo.

Há três etapas para adicionar atalhos de teclado a um complemento:

1. [Configure o manifesto do complemento.](#configure-the-manifest)
1. [Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e seus atalhos de teclado.
1. [Adicione uma ou mais chamadas de tempo de](#create-a-mapping-of-actions-to-their-functions) execução da API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear uma função para cada ação.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Há duas pequenas alterações no manifesto a fazer. Uma é habilitar o add-in para usar um tempo de execução compartilhado e a outra é apontar para um arquivo formatado em JSON onde você definiu os atalhos de teclado.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o complemento para usar um tempo de execução compartilhado

A adição de atalhos de teclado personalizados exige que o seu complemento use o tempo de execução compartilhado. Para obter mais informações, [configure um complemento para usar um tempo de execução compartilhado.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

### <a name="link-the-mapping-file-to-the-manifest"></a>Vincular o arquivo de mapeamento ao manifesto

Imediatamente *abaixo* (não dentro) do `<VersionOverrides>` elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) De definir o atributo para a URL completa de um arquivo JSON em seu projeto `Url` que você criará em uma etapa posterior.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Criar ou editar o arquivo JSON de atalhos

Crie um arquivo JSON em seu projeto. Certifique-se de que o caminho do arquivo corresponde ao local especificado para `Url` o atributo do elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Esse arquivo descreverá os atalhos de teclado e as ações que eles chamarão.

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

    For more information about the JSON objects, see [Constructing the action objects](#constructing-the-action-objects) and [Constructing the shortcut objects](#constructing-the-shortcut-objects). O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Você pode usar "CONTROL" no lugar de "CTRL" neste artigo.

    Em uma etapa posterior, as ações serão mapeadas para as funções que você escrever. Neste exemplo, você mapeará SHOWTASKPANE para uma função que chama o método e HIDETASKPANE para uma função que chama `Office.addin.showAsTaskpane` o `Office.addin.hide` método.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Criar um mapeamento de ações para suas funções

1. No projeto, abra o arquivo JavaScript carregado pela página HTML no `<FunctionFile>` elemento.
1. No arquivo JavaScript, use a API [Office.actions.associate](/javascript/api/office/office.actions#associate) para mapear cada ação especificada no arquivo JSON para uma função JavaScript. Adicione o JavaScript a seguir ao arquivo. Observe o seguinte sobre o código:

    - O primeiro parâmetro é uma das ações do arquivo JSON.
    - O segundo parâmetro é a função que é executado quando um usuário pressiona a combinação de teclas mapeada para a ação no arquivo JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.
1. Para o corpo da função, use o [método Office.addin.showTaskpane](/javascript/api/office/office.addin#showastaskpane--) para abrir o painel de tarefas do complemento. Quando terminar, o código deve se parecer com o seguinte:

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

1. Adicione uma segunda chamada de função para mapear a ação para uma função que chama `Office.actions.associate` `HIDETASKPANE` [Office.addin.hide](/javascript/api/office/office.addin#hide--). Este é um exemplo:

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

Seguir as etapas anteriores permite que o seu complemento alterne a visibilidade do painel de tarefas pressionando **Ctrl+Shift+Tecla** de seta para cima e **Ctrl+Shift+Tecla de seta para baixo.** Esse é o mesmo comportamento mostrado no exemplo de atalhos de teclado [do Excel.](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)

## <a name="details-and-restrictions"></a>Detalhes e restrições

### <a name="constructing-the-action-objects"></a>Construir os objetos de ação

Use as diretrizes a seguir ao especificar os objetos na `action` matriz do shortcuts.jsem:

- Os nomes das `id` propriedades e `name` são obrigatórios.
- A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.
- A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação. Deve ser uma combinação dos caracteres A - Z, a - z, 0 - 9 e as marcas de pontuação "-", "_" e "+".
- A propriedade do `type` é opcional. Atualmente, só `ExecuteFunction` há suporte para o tipo.

Este é um exemplo:

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

O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="constructing-the-shortcut-objects"></a>Construir os objetos de atalho

Use as diretrizes a seguir ao especificar os objetos na `shortcuts` matriz do shortcuts.jsem:

- Os nomes de `action` propriedade , e são `key` `default` obrigatórios.
- O valor da propriedade `action` é uma cadeia de caracteres e deve corresponder a uma das propriedades no objeto `id` action.
- A propriedade pode ser qualquer combinação dos caracteres A - Z, a -z, 0 - 9 e as marcas de `default` pontuação "-", "_" e "+". (Por convenção, letras maiúsculas e baixas não são usadas nessas propriedades.)
- A propriedade deve conter o nome de pelo menos uma tecla modificadora `default` (ALT, CTRL, SHIFT) e apenas uma outra tecla.
- Para Macs, também damos suporte à tecla modificadora COMMAND.
- Para Macs, ALT é mapeada para a tecla OPTION. Para o Windows, COMMAND é mapeado para a tecla CTRL.
- Quando dois caracteres são vinculados à mesma tecla física em um teclado padrão, eles são sinônimos na propriedade; por exemplo, ALT+a e ALT+A são o mesmo atalho, assim como `default` CTRL+- e CTRL+ porque "-" e "_" são a mesma tecla \_ física.
- O caractere "+" indica que as teclas de cada lado são pressionadas simultaneamente.

Este é um exemplo:

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

O esquema completo para os atalhos JSON está [extended-manifest.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Dicas de tecla, também conhecidas como atalhos de tecla sequenciais, como o atalho do Excel para escolher uma cor de preenchimento **Alt+H, H**, não são suportadas nos complementos do Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Usando atalhos quando o foco está no painel de tarefas

Atualmente, os atalhos de teclado para um complemento do Office só podem ser invocados quando o foco do usuário está na planilha. Quando o foco do usuário está dentro da interface do usuário do Office (como o painel de tarefas), nenhum dos atalhos do complemento é ignorado. Como alternativa, o complemento pode definir manipuladores de teclado que podem invocar determinadas ações quando o foco do usuário está dentro da interface do usuário do complemento.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Usando combinações de teclas que já são usadas pelo Office ou outro complemento

Durante o período de visualização, não há sistema para determinar o que acontece quando um usuário pressiona uma combinação de teclas registrada por um complemento e também pelo Office ou por outro. O comportamento é indefinido.

Atualmente, não há uma solução alternativa quando dois ou mais complementos registraram o mesmo atalho de teclado, mas você pode minimizar conflitos com o Excel com estas práticas recomendadas:

- Use somente atalhos de teclado com o seguinte padrão no seu complemento: **Ctrl+Shift+Alt+* x***, onde *x* é alguma outra tecla.
- Se precisar de mais atalhos de teclado, verifique a lista de [atalhos](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)de teclado do Excel e evite usar qualquer um deles no seu complemento.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Atalhos do navegador que não podem ser substituídos

Você não pode usar nenhuma das combinações de teclado a seguir. Eles são usados por navegadores e não podem ser substituídos. Esta lista é um trabalho em andamento. Se você descobrir outras combinações que não podem ser substituídos, nos avise usando a ferramenta de comentários na parte inferior desta página.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="next-steps"></a>Próximas etapas

- Consulte o exemplo de complemento [excel-keyboard-shortcuts](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
