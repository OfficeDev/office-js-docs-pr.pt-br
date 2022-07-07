---
title: Atalhos de teclado personalizados em Suplementos do Office
description: Saiba como adicionar atalhos de teclado personalizados, também conhecidos como combinações de teclas, ao suplemento do Office.
ms.date: 11/22/2021
localization_priority: Normal
ms.openlocfilehash: 5e813e1f4af040bb546f60eb2db40862ba1a237e
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659980"
---
# <a name="add-custom-keyboard-shortcuts-to-your-office-add-ins"></a>Adicionar atalhos de teclado personalizados aos suplementos do Office

Os atalhos de teclado, também conhecidos como combinações de teclas, permitem que os usuários do suplemento trabalhem com mais eficiência. Os atalhos de teclado também melhoram a acessibilidade do suplemento para usuários com deficiências, fornecendo uma alternativa ao mouse.

[!include[Keyboard shortcut prerequisites](../includes/keyboard-shortcuts-prerequisites.md)]

> [!NOTE]
> Para começar com uma versão funcional de um suplemento com atalhos de teclado já habilitados, clone e execute os [atalhos de teclado do Excel de exemplo](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts). Quando estiver pronto para adicionar atalhos de teclado ao seu próprio suplemento, continue com este artigo.

Há três etapas para adicionar atalhos de teclado a um suplemento.

1. [Configure o manifesto do suplemento](#configure-the-manifest).
1. [Crie ou edite o arquivo JSON de atalhos](#create-or-edit-the-shortcuts-json-file) para definir ações e seus atalhos de teclado.
1. [Adicione uma ou mais chamadas de runtime](#create-a-mapping-of-actions-to-their-functions) da API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) para mapear uma função para cada ação.

## <a name="configure-the-manifest"></a>Configurar o manifesto

Há duas pequenas alterações no manifesto a serem feitas. Uma é habilitar o suplemento para usar um runtime compartilhado e a outra é apontar para um arquivo formatado em JSON em que você definiu os atalhos de teclado.

### <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o suplemento para usar um runtime compartilhado

Adicionar atalhos de teclado personalizados exige que o suplemento use o runtime compartilhado. Para obter mais informações, [configure um suplemento para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

### <a name="link-the-mapping-file-to-the-manifest"></a>Vincular o arquivo de mapeamento ao manifesto

Imediatamente *abaixo* (não dentro) do **\<VersionOverrides\>** elemento no manifesto, adicione um [elemento ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Defina `Url` o atributo como a URL completa de um arquivo JSON em seu projeto que você criará em uma etapa posterior.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/shortcuts.json"></ExtendedOverrides>
</OfficeApp>
```

## <a name="create-or-edit-the-shortcuts-json-file"></a>Criar ou editar o arquivo JSON de atalhos

Crie um arquivo JSON em seu projeto. Verifique se o caminho do arquivo corresponde ao local especificado para o `Url` atributo do [elemento ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Esse arquivo descreverá os atalhos de teclado e as ações que eles invocarão.

1. Dentro do arquivo JSON, há duas matrizes. A matriz de ações conterá objetos que definem as ações a serem invocadas e a matriz de atalhos conterá objetos que mapeiam combinações de teclas em ações. Veja um exemplo.
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

    Para obter mais informações sobre os objetos JSON, consulte [Construir os objetos de ação](#construct-the-action-objects) e [construir os objetos de atalho](#construct-the-shortcut-objects). O esquema completo para os atalhos JSON está em [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

    > [!NOTE]
    > Você pode usar "CONTROL" no lugar de "Ctrl" em todo este artigo.

    Em uma etapa posterior, as ações serão mapeadas para as funções que você escreve. Neste exemplo, posteriormente, você mapeará SHOWTASKPANE `Office.addin.showAsTaskpane` para uma função que chama o método e HIDETASKPANE para uma função que chama o `Office.addin.hide` método.

## <a name="create-a-mapping-of-actions-to-their-functions"></a>Criar um mapeamento de ações para suas funções

1. No projeto, abra o arquivo JavaScript carregado pela página HTML no **\<FunctionFile\>** elemento.
1. No arquivo JavaScript, use a API [Office.actions.associate](/javascript/api/office/office.actions#office-office-actions-associate-member) para mapear cada ação especificada no arquivo JSON para uma função JavaScript. Adicione o JavaScript a seguir ao arquivo. Observe o seguinte sobre o código.

    - O primeiro parâmetro é uma das ações do arquivo JSON.
    - O segundo parâmetro é a função que é executada quando um usuário pressiona a combinação de teclas mapeada para a ação no arquivo JSON.

    ```javascript
    Office.actions.associate('-- action ID goes here--', function () {

    });
    ```

1. Para continuar o exemplo, use `'SHOWTASKPANE'` como o primeiro parâmetro.
1. Para o corpo da função, use o [método Office.addin.showAsTaskpane](/javascript/api/office/office.addin#office-office-addin-showastaskpane-member(1)) para abrir o painel de tarefas do suplemento. Quando terminar, o código deverá ser semelhante ao seguinte:

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

1. Adicione uma segunda chamada de `Office.actions.associate` função para mapear `HIDETASKPANE` a ação para uma função que chama [Office.addin.hide](/javascript/api/office/office.addin#office-office-addin-hide-member(1)). Apresentamos um exemplo a seguir.

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

Seguir as etapas anteriores permite que o suplemento alterne a visibilidade do painel de tarefas pressionando **Ctrl+Alt+Para** Cima e **Ctrl+Alt+Para Baixo**. O mesmo comportamento é mostrado no exemplo [de atalhos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) de teclado do Excel no repositório PnP de Suplementos do Office no GitHub.

## <a name="details-and-restrictions"></a>Detalhes e restrições

### <a name="construct-the-action-objects"></a>Construir os objetos de ação

Use as diretrizes a seguir ao especificar os objetos na `actions` matriz do shortcuts.json.

- Os nomes de propriedade `id` e são `name` obrigatórios.
- A `id` propriedade é usada para identificar exclusivamente a ação a ser invocada usando um atalho de teclado.
- A `name` propriedade deve ser uma cadeia de caracteres amigável que descreve a ação. Deve ser uma combinação dos caracteres A - Z, a - z, 0 - 9 e as marcas de pontuação "-", "_" e "+".
- A propriedade do `type` é opcional. Atualmente, há suporte `ExecuteFunction` apenas para o tipo.

Apresentamos um exemplo a seguir.

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

O esquema completo para os atalhos JSON está em [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

### <a name="construct-the-shortcut-objects"></a>Construir os objetos de atalho

Use as diretrizes a seguir ao especificar os objetos na `shortcuts` matriz do shortcuts.json.

- Os nomes de `action`propriedade , `key`e `default` são necessários.
- O valor da propriedade é `action` uma cadeia de caracteres e deve corresponder a uma das `id` propriedades no objeto de ação.
- A `default` propriedade pode ser qualquer combinação dos caracteres A - Z, a -z, 0 - 9 e as marcas de pontuação "-", "_" e "+". (Por convenção, letras minúsculas não são usadas nessas propriedades.)
- A `default` propriedade deve conter o nome de pelo menos uma tecla modificadora (Alt, Ctrl, Shift) e apenas uma outra tecla.
- Shift não pode ser usado como a única tecla modificadora. Combine Shift com Alt ou Ctrl.
- Para Macs, também damos suporte à chave modificadora command.
- Para Macs, Alt é mapeado para a tecla Option. Para o Windows, Command é mapeado para a tecla Ctrl.
- Quando dois caracteres são vinculados à mesma tecla física em um teclado padrão, eles são sinônimos `default` na propriedade; por exemplo, Alt+a e Alt+A são o mesmo atalho, assim como Ctrl+ e Ctrl+\_ porque "-" e "_" são a mesma tecla física.
- O caractere "+" indica que as teclas em ambos os lados são pressionadas simultaneamente.

Apresentamos um exemplo a seguir.

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

O esquema completo para os atalhos JSON está em [extended-manifest.schema.json](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!NOTE]
> Dicas de tecla, também conhecidas como atalhos de tecla sequencial, como o atalho do Excel para escolher uma cor de preenchimento **Alt+H, H**, não têm suporte em Suplementos do Office.

## <a name="avoid-key-combinations-in-use-by-other-add-ins"></a>Evitar combinações de teclas em uso por outros suplementos

Há muitos atalhos de teclado que já estão em uso pelo Office. Evite registrar atalhos de teclado para o suplemento que já estão em uso, no entanto, pode haver algumas instâncias em que é necessário substituir atalhos de teclado existentes ou lidar com conflitos entre vários suplementos que registraram o mesmo atalho de teclado.

No caso de um conflito, o usuário verá uma caixa de diálogo na primeira vez que tentar usar um atalho de teclado conflitantes. Observe que o texto para a `name` opção de suplemento que é exibida nesta caixa de diálogo vem da propriedade no objeto de ação no `shortcuts.json` arquivo.

![Ilustração mostrando um modal de conflito com duas ações diferentes para um único atalho.](../images/add-in-shortcut-conflict-modal.png)

O usuário pode selecionar qual ação o atalho de teclado executará. Depois de fazer a seleção, a preferência é salva para usos futuros do mesmo atalho. As preferências de atalho são salvas por usuário, por plataforma. Se o usuário quiser alterar suas preferências, ele poderá invocar o comando Redefinir preferências de atalho de **Suplementos do Office** na caixa de pesquisa Diga-me. Invocar o comando limpa todas as preferências de atalho do suplemento do usuário e o usuário será solicitado novamente com a caixa de diálogo de conflito na próxima vez que tentar usar um atalho conflitantes.

![A caixa de pesquisa Diga-me no Excel mostrando a ação redefinir as preferências de atalho do Suplemento do Office.](../images/add-in-reset-shortcuts-action.png)

Para obter a melhor experiência do usuário, recomendamos que você minimize conflitos com o Excel com essas boas práticas.

- Use apenas atalhos de teclado com o seguinte padrão: **Ctrl+Shift+Alt+* x***, em que *x* é alguma outra tecla.
- Se precisar de mais atalhos de teclado, verifique a lista de [atalhos](https://support.microsoft.com/office/1798d9d5-842a-42b8-9c99-9b7213f0040f) de teclado do Excel e evite usar qualquer um deles em seu suplemento.
- Quando o foco do teclado estiver dentro da interface do usuário do suplemento, **Ctrl+** Barra de espaços e **Ctrl+Shift+F10** não funcionarão, pois esses são atalhos de acessibilidade essenciais.
- Em um computador Windows ou Mac, se o comando "Redefinir preferências de atalho de Suplementos do Office" não estiver disponível no menu de pesquisa, o usuário poderá adicionar manualmente o comando à faixa de opções personalizando a faixa de opções por meio do menu de contexto.

## <a name="customize-the-keyboard-shortcuts-per-platform"></a>Personalizar os atalhos de teclado por plataforma

É possível personalizar atalhos para serem específicos da plataforma. A seguir está um exemplo do objeto `shortcuts` que personaliza os atalhos para cada uma das seguintes plataformas: `windows`, `mac`, `web`. Observe que você ainda deve ter uma tecla `default` de atalho para cada atalho.

No exemplo a seguir, a `default` chave é a chave de fallback para qualquer plataforma que não seja especificada. A única plataforma não especificada é o Windows, portanto, a `default` chave só será aplicada ao Windows.

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

## <a name="localize-the-keyboard-shortcuts-json"></a>Localizar os atalhos de teclado JSON

Se o suplemento der suporte a várias localidades, você precisará localizar `name` a propriedade dos objetos de ação. Além disso, se qualquer uma das localidades compatíveis com o suplemento tiver alfabetos ou sistemas de escrita diferentes e, portanto, teclados diferentes, talvez seja necessário localizar os atalhos também. Para obter informações sobre como localizar os atalhos de teclado JSON, consulte [Localizar substituições estendidas](../develop/localization.md#localize-extended-overrides).

## <a name="browser-shortcuts-that-cannot-be-overridden"></a>Atalhos do navegador que não podem ser substituídos

Ao usar atalhos de teclado personalizados na Web, alguns atalhos de teclado usados pelo navegador não podem ser substituídos por suplementos. Esta lista é um trabalho em andamento. Se você descobrir outras combinações que não podem ser substituídas, informe-nos usando a ferramenta de comentários na parte inferior desta página.

- Ctrl+N
- Ctrl+Shift+N
- Ctrl+T
- Ctrl+Shift+T
- Ctrl+W
- Ctrl+PgUp/PgDn

## <a name="enable-custom-keyboard-shortcuts-for-specific-users"></a>Habilitar atalhos de teclado personalizados para usuários específicos

O suplemento pode permitir que os usuários reatribuam as ações do suplemento a combinações de teclado alternativas.

> [!NOTE]
> As APIs descritas nesta seção exigem o conjunto de [requisitos keyboardShortcuts 1.1](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets) .

Use o [método Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) para atribuir combinações de teclado personalizadas de um usuário às suas ações de suplementos. O método usa um `{[actionId:string]: string|null}`parâmetro do tipo, `actionId`em que os s são um subconjunto das IDs de ação que devem ser definidas no JSON de manifesto estendido do suplemento. Os valores são as combinações de teclas preferenciais do usuário. O valor também pode ser `null`, `actionId` que removerá qualquer personalização para isso e reverterá para a combinação de teclado padrão definida no JSON de manifesto estendido do suplemento.

Se o usuário estiver conectado ao Office, as combinações personalizadas serão salvas nas configurações de roaming do usuário por plataforma. Atualmente, não há suporte para a personalização de atalhos para usuários anônimos.

```javascript
const userCustomShortcuts = {
    SHOWTASKPANE:"CTRL+SHIFT+1", 
    HIDETASKPANE:"CTRL+SHIFT+2"
};
Office.actions.replaceShortcuts(userCustomShortcuts)
    .then(function () {
        console.log("Successfully registered.");
    })
    .catch(function (ex) {
        if (ex.code == "InvalidOperation") {
            console.log("ActionId does not exist or shortcut combination is invalid.");
        }
    });
```

Para descobrir quais atalhos já estão em uso para o usuário, chame o [método Office.actions.getShortcuts](/javascript/api/office/office.actions#office-office-actions-getshortcuts-member) . Esse método retorna um objeto do tipo `[actionId:string]:string|null}`, em que os valores representam a combinação de teclado atual que o usuário deve usar para invocar a ação especificada. Os valores podem vir de três fontes diferentes:

- Se houver um conflito com o atalho e o usuário tiver optado por usar uma ação diferente (nativa ou outro suplemento) para essa combinação de teclado, `null` o valor retornado será uma vez que o atalho foi substituído e não há nenhuma combinação de teclado que o usuário possa usar no momento para invocar essa ação de suplemento.
- Se o atalho tiver sido personalizado usando o método [Office.actions.replaceShortcuts](/javascript/api/office/office.actions#office-office-actions-replaceshortcuts-member) , o valor retornado será a combinação de teclado personalizada.
- Se o atalho não tiver sido substituído ou personalizado, ele retornará o valor do JSON do manifesto estendido do suplemento.

Apresentamos um exemplo a seguir.

```javascript
Office.actions.getShortcuts()
    .then(function (userShortcuts) {
       for (const action in userShortcuts) {
           let shortcut = userShortcuts[action];
           console.log(action + ": " + shortcut);
       }
    });

```

Conforme descrito [em Evitar combinações de teclas](#avoid-key-combinations-in-use-by-other-add-ins) em uso por outros suplementos, é uma boa prática evitar conflitos em atalhos. Para descobrir se uma ou mais combinações de teclas já estão em uso, passe-as como uma matriz de cadeias de caracteres para o método [Office.actions.areShortcutsInUse](/javascript/api/office/office.actions#office-office-actions-areshortcutsinuse-member) . O método retorna um relatório que contém combinações de teclas que já estão em uso na forma de uma matriz de objetos do tipo `{shortcut: string, inUse: boolean}`. A `shortcut` propriedade é uma combinação de teclas, como "CTRL+SHIFT+1". Se a combinação já estiver registrada em outra ação, a `inUse` propriedade será definida como `true`. Por exemplo, `[{shortcut: "CTRL+SHIFT+1", inUse: true}, {shortcut: "CTRL+SHIFT+2", inUse: false}]`. O snippet de código a seguir é um exemplo:

```javascript
const shortcuts = ["CTRL+SHIFT+1", "CTRL+SHIFT+2"];
Office.actions.areShortcutsInUse(shortcuts)
    .then(function (inUseArray) {
        const availableShortcuts = inUseArray.filter(function (shortcut) { return !shortcut.inUse; });
        console.log(availableShortcuts);
        const usedShortcuts = inUseArray.filter(function (shortcut) { return shortcut.inUse; });
        console.log(usedShortcuts);
    });

```

## <a name="next-steps"></a>Próximas etapas

- Consulte o [suplemento de exemplo de atalhos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) de teclado do Excel.
- Obtenha uma visão geral de como trabalhar com substituições estendidas no [Work com substituições estendidas do manifesto](../develop/extended-overrides.md).
