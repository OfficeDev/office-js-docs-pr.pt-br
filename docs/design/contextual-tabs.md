---
title: Criar guias contextuais personalizadas em Complementos do Office
description: Saiba como adicionar guias contextuais personalizadas ao seu Complemento do Office.
ms.date: 01/20/2021
localization_priority: Normal
ms.openlocfilehash: 7c9593c98bf7cc7f4e270037768be1e2de06aeb3
ms.sourcegitcommit: 1d33ea6dd3a55fd3bc9af48737ad6d7369d30cd8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/22/2021
ms.locfileid: "49934342"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)

Uma guia contextual é um controle guia oculto na faixa de opções do Office que é exibido na linha da guia quando um evento especificado ocorre no documento do Office. Por exemplo, a **guia Design de** Tabela que aparece na faixa de opções do Excel quando uma tabela é selecionada. Você pode incluir guias contextuais personalizadas no seu complemento do Office e especificar quando elas ficam visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade. (No entanto, as guias contextuais personalizadas não respondem a alterações de foco.)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

> [!IMPORTANT]
> As guias contextuais personalizadas estão em visualização. Experimente-os em um ambiente de desenvolvimento ou teste, mas não os adicione a um complemento de produção.
>
> Atualmente, as guias contextuais personalizadas só têm suporte no Excel e apenas nessas plataformas e builds:
>
> - Excel no Windows (somente Microsoft 365, não licença permanente): Versão 2011 (Build 13426.20274). Sua assinatura do Microsoft 365 pode precisar estar no Canal Atual [(Visualização)](https://insider.office.com/join/windows) anteriormente chamado de "Canal Mensal (Direcionado)" ou "Participante do Insider - Lento".

> [!NOTE]
> As guias contextuais personalizadas funcionam apenas em plataformas que suportam os seguintes conjuntos de requisitos. Para saber mais sobre conjuntos de requisitos e como trabalhar com eles, confira [Especificar aplicativos do Office e requisitos de API.](../develop/specify-office-hosts-and-api-requirements.md)
>
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>Comportamento de guias contextuais personalizadas

A experiência do usuário para guias contextuais personalizadas segue o padrão das guias contextuais internas do Office. A seguir estão os princípios básicos para as guias contextuais personalizadas de posicionamento:

- Quando uma guia contextual personalizada está visível, ela aparece na extremidade direita da faixa de opções.
- Se uma ou mais guias contextuais e uma ou mais guias contextuais personalizadas de complementos estão visíveis ao mesmo tempo, as guias contextuais personalizadas estão sempre à direita de todas as guias contextuais.
- Se o seu add-in tiver mais de uma guia contextual e houver contextos nos quais mais de uma está visível, eles aparecerão na ordem em que estão definidos no seu complemento. (A direção tem a mesma direção do idioma do Office, ou seja, da esquerda para a direita, nos idiomas da esquerda para a direita, mas da direita para a esquerda nos idiomas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como defini-los.
- Se mais de um complemento tiver uma guia contextual visível em um contexto específico, elas aparecerão na ordem em que os complementos foram lançados.
- As *guias contextuais* personalizadas, ao contrário das guias principais personalizadas, não são adicionadas permanentemente à faixa de opções do aplicativo do Office. Eles estão presentes somente em documentos do Office nos quais o seu complemento está sendo executado.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Principais etapas para incluir uma guia contextual em um complemento

Veja a seguir as principais etapas para incluir uma guia contextual personalizada em um complemento:

1. Configure o complemento para usar um tempo de execução compartilhado.
1. Defina a guia e os grupos e controles que aparecem nele.
1. Registre a guia contextual no Office.
1. Especifique as circunstâncias em que a guia ficará visível.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o complemento para usar um tempo de execução compartilhado

A adição de guias contextuais personalizadas exige que o seu complemento use o tempo de execução compartilhado. Para obter mais informações, [consulte Configurar um complemento para usar um tempo de execução compartilhado.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Definir os grupos e controles que aparecem na guia

Ao contrário das guias principais personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas em tempo de execução com um blob JSON. Seu código analisará o blob em um objeto JavaScript e passará o objeto para o [método Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) Guias contextuais personalizadas só estão presentes em documentos nos quais seu complemento está sendo executado no momento. Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções do aplicativo do Office quando o complemento é instalado e permanecem presentes quando outro documento é aberto. Além disso, `requestCreateControls` o método pode ser executado apenas uma vez em uma sessão do seu complemento. Se for chamado novamente, será lançado um erro.

> [!NOTE]
> A estrutura das propriedades e subpropriedades do blob JSON (e os nomes de chave) é aproximadamente paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no manifesto XML.

Vamos construir um exemplo de um blob JSON de guias contextuais passo a passo. (O esquema completo para a guia contextual JSON está [dynamic-ribbon.schema.jsem](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Este link pode não estar funcionando no período de visualização antecipada para guias contextuais. Se o link não estiver funcionando, você poderá encontrar o rascunho mais recente do esquema em rascunho [dynamic-ribbon.schema.jsem](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Se você estiver trabalhando no Visual Studio Code, poderá usar esse arquivo para obter o IntelliSense e validar seu JSON. Para obter mais informações, consulte Edição JSON com o Visual Studio Code - esquemas [e configurações JSON.](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings)


1. Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz nomeadas `actions` e `tabs` . A matriz é uma especificação de todas as funções que podem ser executadas por `actions` controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, até um máximo de *20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação. Adicione o seguinte como o único membro da `actions` matriz. Sobre essa marcação, observe:

    - As `id` propriedades e as propriedades são `type` obrigatórias.
    - O valor pode `type` ser "ExecuteFunction" ou "ShowTaskpane".
    - A `functionName` propriedade só é usada quando o valor é `type` `ExecuteFunction` . É o nome de uma função definida no FunctionFile. Para obter mais informações sobre o FunctionFile, consulte [Conceitos básicos para comandos de complemento.](add-in-commands.md)
    - Em uma etapa posterior, você mapeará essa ação para um botão na guia contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Adicione o seguinte como o único membro da `tabs` matriz. Sobre essa marcação, observe:

    - A propriedade `id` é obrigatória. Use uma ID breve e descritiva que seja exclusiva entre todas as guias contextuais do seu complemento.
    - A propriedade `label` é obrigatória. É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.
    - A propriedade `groups` é obrigatória. Ele define os grupos de controles que aparecerão na guia. Ele deve ter pelo menos um membro *e não mais de 20.* (Há também limites no número de controles que você pode ter em uma guia contextual personalizada e que também restringirá quantos grupos você tem. Consulte a próxima etapa para obter mais informações.)

    > [!NOTE]
    > O objeto tab também pode ter uma propriedade opcional que especifica se a guia é visível `visible` imediatamente quando o complemento é iniciado. Como as guias contextuais normalmente ficam ocultas até que um evento do usuário acione sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a propriedade assume como padrão quando não está `visible` `false` presente. Em uma seção posterior, mostraremos como definir a propriedade em `true` resposta a um evento.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. No exemplo contínuo simples, a guia contextual tem apenas um único grupo. Adicione o seguinte como o único membro da `groups` matriz. Sobre essa marcação, observe:

    - Todas as propriedades são necessárias.
    - A `id` propriedade deve ser exclusiva entre todos os grupos na guia. Use uma ID breve e descritiva.
    - É `label` uma cadeia de caracteres amigável para servir como o rótulo do grupo.
    - O valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela do aplicativo `icon` do Office.
    - O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e menus no grupo. Deve haver pelo menos um.

    > [!IMPORTANT]
    > *O número total de controles na guia inteira não pode ser maior que 20.* Por exemplo, você pode ter 3 grupos com 6 controles cada e um quarto grupo com 2 controles, mas não pode ter 4 grupos com 6 controles cada.  

    ```json
    {
        "id": "CustomGroup111",
        "label": "Insertion",
        "icon": [

        ],
        "controls": [

        ]
    }
    ```

1. Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32 x 32 px e 80x80 px. Opcionalmente, você também pode ter ícones de tamanhos de 16 x 16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px. O Office decide qual ícone usar com base no tamanho da faixa de opções e da janela do aplicativo do Office. Adicione os seguintes objetos à matriz de ícones. (Se os tamanhos da janela e da  faixa de opções são grandes o suficiente para que pelo menos um dos controles do grupo apareça, nenhum ícone de grupo é exibido. Por exemplo, assista ao grupo **Estilos** na faixa de opções do Word conforme você reduz e expande a janela do Word.) Sobre essa marcação, observe:

    - Ambas as propriedades são necessárias.
    - A `size` unidade de medida da propriedade é pixels. Os ícones são sempre quadrados, portanto, o número é a altura e a largura.
    - A `sourceLocation` propriedade especifica a URL completa para o ícone.

    > [!IMPORTANT]
    > Assim como normalmente você deve alterar as URLs no manifesto do add-in quando você muda do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.

    ```json
    {
        "size": 32,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
    },
    {
        "size": 80,
        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
    }
    ```

1. No nosso exemplo contínuo simples, o grupo tem apenas um único botão. Adicione o seguinte objeto como o único membro da `controls` matriz. Sobre essa marcação, observe:

    - Todas as propriedades, exceto `enabled` , são necessárias.
    - `type` especifica o tipo de controle. Os valores podem ser "Button", "Menu" ou "MobileButton".
    - `id` pode ter até 125 caracteres. 
    - `actionId` deve ser a ID de uma ação definida na `actions` matriz. (Consulte a etapa 1 desta seção.)
    - `label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.
    - `superTip` representa uma forma rica de dica de ferramenta. As propriedades `title` e as propriedades são `description` necessárias.
    - `icon` especifica os ícones do botão. Os comentários anteriores sobre o ícone de grupo também se aplicam aqui.
    - `enabled` (opcional) especifica se o botão está habilitado quando a guia contextual aparece iniciando. O padrão se não estiver presente é `true` . 

    ```json
    {
        "type": "Button",
        "id": "CtxBt112",
        "actionId": "executeWriteData",
        "enabled": false,
        "label": "Write Data",
        "superTip": {
            "title": "Data Insertion",
            "description": "Use this button to insert data into the document."
        },
        "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
            }
        ]
    }
    ```
 
Veja a seguir o exemplo completo do blob JSON:

```json
`{
  "actions": [
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
  ],
  "tabs": [
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [
        {
          "id": "CustomGroup111",
          "label": "Insertion",
          "icon": [
            {
                "size": 32,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group32x32.png"
            },
            {
                "size": 80,
                "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/Group80x80.png"
            }
          ],
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "executeWriteData",
                "enabled": false,
                "label": "Write Data",
                "superTip": {
                    "title": "Data Insertion",
                    "description": "Use this button to insert data into the document."
                },
                "icon": [
                    {
                        "size": 32,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton32x32.png"
                    },
                    {
                        "size": 80,
                        "sourceLocation": "https://cdn.contoso.com/addins/datainsertion/Images/WriteDataButton80x80.png"
                    }
                ]
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Registrar a guia contextual com o Office com requestCreateControls

A guia contextual é registrada com o Office chamando o [método Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Isso geralmente é feito na função atribuída a `Office.initialize` ou com o `Office.onReady` método. Para saber mais sobre esses métodos e como inicializar o add-in, confira Inicializar seu [complemento do Office.](../develop/initialize-add-in.md) No entanto, você pode chamar o método a qualquer momento após a inicialização.

> [!IMPORTANT]
> O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um complemento. Um erro será lançado se for chamado novamente.

Apresentamos um exemplo a seguir. Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o método antes que ela possa ser passada para `JSON.parse` uma função JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Especificar os contextos quando a guia ficará visível com requestUpdate

Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto do complemento. Considere um cenário no qual a guia deve estar visível quando, e somente quando, um gráfico (na planilha padrão de uma pasta de trabalho do Excel) é ativado.

Comece atribuindo manipuladores. Isso geralmente é feito no método como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos eventos de todos os `Office.onReady` gráficos `onActivated` na `onDeactivated` planilha.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Em seguida, defina os manipuladores. Veja a seguir um exemplo simples de um erro , mas consulte Manipulando o erro `showDataTab` [HostRestartNeeded](#handling-the-hostrestartneeded-error) posteriormente neste artigo para obter uma versão mais robusta da função. Sobre este código, observe:

- O Office controla quando atualiza o estado da faixa de opções. O  [método Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) enfiltrou uma solicitação para atualizar. O método resolverá o objeto assim que a solicitação estiver na fila, não quando a faixa de opções `Promise` for realmente atualizada.
- O parâmetro para o método é um objeto `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia por sua ID exatamente como especificado no *JSON* e (2) especifica a visibilidade da guia.
- Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar outros objetos tab à `tabs` matriz.

```javascript
async function showDataTab() {
    await Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            }
        ]});
}
```

O manipulador para ocultar a guia é quase idêntico, exceto pelo fato de que ele define `visible` a propriedade novamente como `false` .

A biblioteca JavaScript do Office também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto. A seguir está `showDataTab` a função em TypeScript e ela faz uso desses tipos.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Visibilidade da guia de alternância e o status habilitado de um botão ao mesmo tempo

O método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada ou `requestUpdate` em uma guia principal personalizada. Para obter detalhes sobre isso, [consulte Habilitar e desabilitar comandos de complemento.](disable-add-in-commands.md) Pode haver cenários em que você queira alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo. Você pode fazer isso com uma única chamada de `requestUpdate` . A seguir está um exemplo no qual um botão em uma guia principal é habilitado ao mesmo tempo que uma guia contextual é visível.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true
            },
            {
                id: "OfficeAppTab1",
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                            }
                        ]
                    }
                ]
            ]}
        ]
    });
}
```

No exemplo a seguir, o botão que está habilitado está na mesma guia contextual que está sendo visível.

```javascript
function myContextChanges() {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "CtxTab1",
                visible: true,
                groups: [
                    {
                        id: "CustomGroup111",
                        controls: [
                            {
                                id: "MyButton",
                                enabled: true
                           }
                       ]
                   }
               ]
            }
        ]
    });
}
```

## <a name="localizing-the-json-blob"></a>Localizando o blob JSON

O blob JSON que é passado não é localizado da mesma maneira que a marcação de manifesto para guias principais personalizadas é localizada (que é descrito na localização de controle do `requestCreateControls` [manifesto](../develop/localization.md#control-localization-from-the-manifest)). Em vez disso, a localização deve ocorrer em tempo de execução usando blobs JSON distintos para cada localidade. Sugerimos que você use uma instrução que teste a `switch` [propriedade Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Veja um exemplo a seguir:

```javascript
function GetContextualTabsJsonSupportedLocale () {
    var displayLanguage = Office.context.displayLanguage;

        switch (displayLanguage) {
            case 'en-US':
                return `{
                    "actions": [
                        // actions omitted
                     ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Data",
                          "groups": [
                              // groups omitted
                          ]
                        }
                    ]
                }`;

            case 'fr-FR':
                return `{
                    "actions": [
                        // actions omitted 
                    ],
                    "tabs": [
                        {
                          "id": "CtxTab1",
                          "label": "Contoso Données",
                          "groups": [
                              // groups omitted
                          ]
                       }
                    ]
               }`;

            // Other cases omitted
       }
}
```

Em seguida, seu código chama a função para obter o blob localizado que é passado `requestCreateControls` para, como no exemplo a seguir:

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="handling-the-hostrestartneeded-error"></a>Manipulando o erro HostRestartNeeded

Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro. Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto. Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`. Veja um exemplo de como lidar com esse erro a seguir. Nesse caso, o método `reportError` exibe o erro para o usuário.

```javascript
function showDataTab() {
    try {
        await Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "CtxTab1",
                    visible: true
                }
            ]});
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, then close and reopen the Office application.");
        }
    }
}
```
