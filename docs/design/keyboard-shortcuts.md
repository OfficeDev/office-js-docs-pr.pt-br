---
title: Atalhos de teclado personalizados em suplementos do Office
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao suplemento do Office.
ms.date: 11/09/2020
localization_priority: Normal
ms.openlocfilehash: f95c26067203a4ec2659aa6a632403c96ed81674
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996671"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins-preview"></a>Adicionar atalhos de teclado personalizados para seus suplementos do Office (visualização)

Os atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do seu suplemento trabalhem com mais eficiência e melhoram a acessibilidade do suplemento aos usuários com deficiências, fornecendo uma alternativa ao mouse.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Para começar com uma versão de trabalho de um suplemento com atalhos de teclado já habilitados, clone e execute os [atalhos de teclado do Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts)de exemplo. Quando estiver pronto para adicionar atalhos de teclado ao seu próprio suplemento, continue com este artigo.

Há três etapas para adicionar atalhos de teclado a um suplemento:

1. [Configure o manifesto do suplemento](#configure-the-manifest).
1. [Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e seus atalhos de teclado.
1. [Adicione uma ou mais chamadas de tempo de execução](#create-a-mapping-of-actions-to-their-functions) da API [Office. Actions. associa](/javascript/api/office/office.actions#associate) para mapear uma função para cada ação.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Há duas pequenas alterações para o manifesto fazer. Uma é habilitar o suplemento para usar um tempo de execução compartilhado e o outro é apontar para um arquivo formatado por JSON onde você definiu os atalhos de teclado.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o suplemento para usar um tempo de execução compartilhado

Adicionar atalhos de teclado personalizados exige que seu suplemento use o tempo de execução compartilhado. Para obter mais informações, [Configure um suplemento para usar um tempo de execução compartilhado](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Vincular o arquivo de mapeamento ao manifesto

Imediatamente *abaixo* (não dentro) do `<VersionOverrides>` elemento no manifesto, adicione um elemento [ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Defina o `Url` atributo para a URL completa de um arquivo JSON em seu projeto que você irá criar em uma etapa posterior.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Criar ou editar o arquivo JSON de atalhos

Crie um arquivo JSON em seu projeto. Certifique-se de que o caminho do arquivo corresponde ao local especificado para o `Url` atributo do elemento [ExtendedOverrides](../reference/manifest/extendedoverrides.md) . Esse arquivo descreve seus atalhos de teclado e as ações que eles invocarão.

1. Dentro do arquivo JSON, há duas matrizes. A matriz de ações conterá objetos que definem as ações a serem invocadas e a matriz de atalhos conterá objetos que mapeiam combinações de teclas para ações. Veja um exemplo:

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

    Para obter mais informações sobre os objetos JSON, consulte [construir os objetos Action](#constructing-the-action-objects) e [criar os objetos de atalho](#constructing-the-shortcut-objects). O esquema completo para o JSON de atalhos está em [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Você pode usar o "controle" em vez de "CTRL" neste artigo.

    Em uma etapa posterior, as ações serão mapeadas para as funções que você escrever. Neste exemplo, posteriormente, você irá mapear SHOWTASKPANE para uma função que chama o `Office.addin.showAsTaskpane` método e HIDETASKPANE para uma função que chama o `Office.addin.hide` método.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Criar um mapeamento de ações para suas funções

1. Em seu projeto, abra o arquivo JavaScript carregado pela página HTML no `<FunctionFile>` elemento.
1. No arquivo JavaScript, use a API [Office. Actions. associa](/javascript/api/office/office.actions#associate) para mapear cada ação que você especificou no arquivo JSON para uma função JavaScript. Adicione o seguinte JavaScript ao arquivo. Observe o seguinte sobre o código:

    - O primeiro parâmetro é uma das ações do arquivo JSON.
    - O segundo parâmetro é a função que é executada quando um usuário pressiona a combinação de teclas que é mapeada para a ação no arquivo JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.
1. Para o corpo da função, use o método [Office. AddIn. showTaskpane](/javascript/api/office/office.addin.md#showastaskpane--) para abrir o painel de tarefas do suplemento. Quando você terminar, o código deverá ser semelhante ao seguinte:

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

1. Adicione uma segunda chamada de `Office.actions.associate` função para mapear a `HIDETASKPANE` ação para uma função que chama [Office. AddIn. Hide](/javascript/api/office/office.addin.md#hide--). Este é um exemplo:

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

Seguindo as etapas anteriores permite que o suplemento alterne a visibilidade do painel de tarefas pressionando **Ctrl + Shift + tecla de seta para cima** e **Ctrl + Shift + tecla de seta para baixo**. Esse é o mesmo comportamento mostrado no suplemento de [exemplo de atalhos de teclado do Excel](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).

## <a name="details-and-restrictions"></a>Detalhes e restrições

### <a name="constructing-the-action-objects"></a>Construir os objetos Action

Use as diretrizes a seguir ao especificar os objetos na `action` matriz de shortcuts.jsem:

- Os nomes das propriedades `id` e `name` são obrigatórios.
- A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.
- A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação. Deve ser uma combinação dos caracteres A-Z, a-z, 0-9 e das marcas de Pontuação "-", "_" e "+".
- A propriedade do `type` é opcional. No momento, só `ExecuteFunction` há suporte para Type.

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

O esquema completo para o JSON de atalhos está em [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

### <a name="constructing-the-shortcut-objects"></a>Construir os objetos de atalho

Use as diretrizes a seguir ao especificar os objetos na `shortcuts` matriz de shortcuts.jsem:

- Os nomes das propriedades `action` , `key` e `default` são obrigatórios.
- O valor da `action` propriedade é uma cadeia de caracteres e deve corresponder a uma das `id` Propriedades no objeto Action.
- A `default` propriedade pode ser qualquer combinação dos caracteres a-z, a-z, 0-9 e das marcas de Pontuação "-", "_" e "+". (Por convenção, letras minúsculas não são usadas nessas propriedades.)
- A `default` propriedade deve conter o nome de pelo menos uma tecla modificadora (Alt, CTRL, Shift) e apenas uma tecla.
- Para Macs, também há suporte para a tecla modificador de comandos.
- Para Macs, ALT é mapeada para a tecla de opção. Para o Windows, o comando é mapeado para a tecla CTRL.
- Quando dois caracteres são vinculados à mesma chave física em um teclado padrão, eles são sinônimos na `default` Propriedade; por exemplo, ALT + a e Alt + a são o mesmo atalho, portanto, são CTRL + e CTRL + \_ porque "-" e "_" são a mesma chave física.
- O caractere "+" indica que as teclas de ambos os lados são pressionadas simultaneamente.

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

O esquema completo para o JSON de atalhos está em [extended-manifest.schema.js](https://developer.microsoft.com/en-us/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> As dicas de teclas, também conhecidas como atalhos de chave sequencial, como o atalho do Excel para escolher uma cor de preenchimento **ALT + H** , não são compatíveis com os suplementos do Office.

### <a name="using-shortcuts-when-the-focus-is-in-the-task-pane"></a>Usando atalhos quando o foco está no painel de tarefas

Atualmente, os atalhos de teclado para um suplemento do Office só podem ser invocados quando o foco do usuário está na planilha. Quando o foco do usuário está dentro da interface do usuário do Office (como o painel de tarefas), nenhum dos atalhos do suplemento é ignorado. Como uma solução alternativa, o suplemento pode definir manipuladores de teclado que podem invocar determinadas ações quando o foco do usuário está dentro da interface do usuário do suplemento.

## <a name="using-key-combinations-that-are-already-used-by-office-or-another-add-in"></a>Usando combinações de teclas já utilizadas pelo Office ou outro suplemento

Durante o período de visualização, não há nenhum sistema para determinar o que acontece quando um usuário pressiona uma combinação de teclas que é registrada por um suplemento e também pelo Office ou por outro suplemento. O comportamento é indefinido.

No momento, não há solução alternativa quando dois ou mais suplementos registraram o mesmo atalho de teclado, mas você pode minimizar conflitos com o Excel com essas boas práticas:

- Use apenas atalhos de teclado com o seguinte padrão em seu suplemento: * *Ctrl + Shift + Alt +* x * * *, onde *x* é outra tecla.
- Se você precisar de mais atalhos de teclado, verifique a [lista de atalhos de teclado do Excel](https://support.microsoft.com/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f)e evite usar qualquer um deles no suplemento.

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Atalhos do navegador que não podem ser substituídos

Você não pode usar nenhuma das combinações de teclado a seguir. Eles são usados pelos navegadores e não podem ser substituídos. Esta lista é um trabalho em andamento. Se você descobrir outras combinações que não podem ser substituídas, informe-nos usando a ferramenta de comentários na parte inferior desta página.

- Ctrl + N
- Ctrl + Shift + N
- CTRL + T
- CTRL + SHIFT + T
- CTRL + W
- CTRL + PgUp/PgDn

## <a name="next-steps"></a>Próximas Etapas

- Confira o suplemento de exemplo [Excel-teclado-atalhos](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-keyboard-shortcuts).
