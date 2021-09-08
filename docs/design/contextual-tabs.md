---
title: Criar guias contextuais personalizadas em Office de complementos
description: Saiba como adicionar guias contextuais personalizadas ao seu Office Add-in.
ms.date: 09/02/2021
localization_priority: Normal
ms.openlocfilehash: 3efcc29ea78d7dd528734e2c67a14cd65e3c0875
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938204"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Criar guias contextuais personalizadas em Office de complementos

Uma guia contextual é um controle de tabulação oculto na faixa Office que é exibido na linha de tabulação quando um evento especificado ocorre no documento Office. Por exemplo, a **guia Design** de Tabela que aparece na faixa Excel quando uma tabela é selecionada. Você inclui guias contextuais personalizadas no seu Office e especifica quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade. (No entanto, as guias contextuais personalizadas não respondem a alterações de foco.)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

[!INCLUDE [Animation of contextual tabs and enabling buttons](../includes/animation-contextual-tabs-enable-button.md)]

> [!IMPORTANT]
> Atualmente, as guias contextuais personalizadas só têm suporte Excel e somente nessas plataformas e builds.
>
> - Excel no Windows (somente Microsoft 365 assinatura): Versão 2102 (Build 13801.20294) ou posterior.
> - Excel Online

> [!NOTE]
> As guias contextuais personalizadas funcionam somente em plataformas que suportam os seguintes conjuntos de requisitos. Para obter mais informações sobre conjuntos de requisitos e como trabalhar com eles, consulte [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Você pode usar as verificações de tempo de execução em seu código para testar se a combinação de host e plataforma do usuário oferece suporte a esses conjuntos de requisitos, conforme descrito em Especificar aplicativos Office e requisitos [de API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). (A técnica de especificar os conjuntos de requisitos no manifesto, que também é descrito nesse artigo, não funciona atualmente para RibbonApi 1.2.) Como alternativa, você pode [implementar uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportamento de guias contextuais personalizadas

A experiência do usuário para guias contextuais personalizadas segue o padrão de guias Office contextuais internas. A seguir estão os princípios básicos para as guias contextuais personalizadas de posicionamento.

- Quando uma guia contextual personalizada é visível, ela aparece na extremidade direita da faixa de opções.
- Se uma ou mais guias contextuais e uma ou mais guias contextuais personalizadas de complementos são visíveis ao mesmo tempo, as guias contextuais personalizadas estão sempre à direita de todas as guias contextuais.
- Se o seu add-in tiver mais de uma guia contextual e houver contextos nos quais mais de um está visível, eles aparecerão na ordem em que são definidos no seu complemento. (A direção é a mesma direção que o idioma Office, ou seja, é da esquerda para a direita em idiomas da esquerda para a direita, mas da direita para a esquerda em idiomas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como defini-los.
- Se mais de um complemento tiver uma guia contextual visível em um contexto específico, elas aparecerão na ordem na qual os complementos foram lançados.
- Guias *contextuais personalizadas,* ao contrário das guias principais personalizadas, não são adicionadas permanentemente à faixa Office do aplicativo. Eles estão presentes apenas em Office documentos nos quais o seu complemento está sendo executado.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Principais etapas para incluir uma guia contextual em um complemento

A seguir estão as principais etapas para incluir uma guia contextual personalizada em um complemento.

1. Configure o complemento para usar um tempo de execução compartilhado.
1. Defina a guia e os grupos e controles que aparecem nele.
1. Registre a guia contextual com Office.
1. Especifique as circunstâncias em que a guia ficará visível.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o add-in para usar um tempo de execução compartilhado

A adição de guias contextuais personalizadas exige que o seu add-in use o tempo de execução compartilhado. Para obter mais informações, [consulte Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Definir os grupos e controles que aparecem na guia

Ao contrário das guias principais personalizadas, definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com um blob JSON. Seu código analisará o blob em um objeto JavaScript e passará o objeto para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) As guias contextuais personalizadas só estão presentes em documentos nos quais o seu complemento está sendo executado no momento. Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções do aplicativo Office quando o complemento é instalado e permanecem presentes quando outro documento é aberto. Além disso, `requestCreateControls` o método pode ser executado apenas uma vez em uma sessão do seu complemento. Se for chamado novamente, será lançado um erro.

> [!NOTE]
> A estrutura das propriedades e subpropropriedades do blob JSON (e os nomes principais) é aproximadamente paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no XML do manifesto.

Construiremos um exemplo de guias contextuais blob JSON passo a passo. O esquema completo da guia contextual JSON estádynamic-ribbon.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Se você estiver trabalhando no Visual Studio Code, poderá usar esse arquivo para obter IntelliSense e validar seu JSON. Para obter mais informações, [consulte Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz nomeadas `actions` e `tabs` . A `actions` matriz é uma especificação de todas as funções que podem ser executadas por controles na guia contextual. A matriz define uma ou mais guias `tabs` contextuais, *até um máximo de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação. Adicione o seguinte como o único membro da `actions` matriz. Sobre essa marcação, observe:

    - As `id` propriedades e são `type` obrigatórias.
    - O valor de `type` pode ser "ExecuteFunction" ou "ShowTaskpane".
    - A `functionName` propriedade só é usada quando o valor de é `type` `ExecuteFunction` . É o nome de uma função definida no FunctionFile. Para obter mais informações sobre o FunctionFile, consulte [Conceitos básicos para Comandos de Complemento.](add-in-commands.md)
    - Em uma etapa posterior, você mapeará essa ação para um botão na guia contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Adicione o seguinte como o único membro da `tabs` matriz. Sobre essa marcação, observe:

    - A propriedade `id` é obrigatória. Use uma ID breve e descritiva que seja exclusiva entre todas as guias contextuais no seu complemento.
    - A propriedade `label` é obrigatória. É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.
    - A propriedade `groups` é obrigatória. Ele define os grupos de controles que serão exibidos na guia. Ele deve ter pelo menos um membro *e não mais de 20*. (Também há limites sobre o número de controles que você pode ter em uma guia contextual personalizada e que também restringirá quantos grupos você tem. Consulte a próxima etapa para obter mais informações.)

    > [!NOTE]
    > O objeto tab também pode ter uma propriedade opcional que especifica se a guia fica `visible` visível imediatamente quando o complemento é iniciado. Como as guias contextuais normalmente ficam ocultas até que um evento do usuário acione sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a propriedade é padrão quando não está `visible` `false` presente. Em uma seção posterior, mostramos como definir a propriedade como `true` em resposta a um evento.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. No exemplo simples em andamento, a guia contextual tem apenas um único grupo. Adicione o seguinte como o único membro da `groups` matriz. Sobre essa marcação, observe:

    - Todas as propriedades são necessárias.
    - A `id` propriedade deve ser exclusiva entre todos os grupos na guia. Use uma ID breve e descritiva.
    - A `label` é uma cadeia de caracteres amigável para servir como o rótulo do grupo.
    - O valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela Office `icon` aplicativo.
    - O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e os menus no grupo. Deve haver pelo menos um.

    > [!IMPORTANT]
    > *O número total de controles na guia inteira não pode ser maior do que 20.* Por exemplo, você pode ter 3 grupos com 6 controles cada e um quarto grupo com 2 controles, mas não pode ter 4 grupos com 6 controles cada.  

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

1. Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32x32 px e 80x80 px. Opcionalmente, você também pode ter ícones de tamanhos 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px. Office decide qual ícone usar com base no tamanho da faixa de opções e Office de aplicativos. Adicione os seguintes objetos à matriz de ícones. (Se os tamanhos da janela e da  faixa de opções são grandes o suficiente para que pelo menos um dos controles no grupo apareça, nenhum ícone de grupo será exibido. Por exemplo, assista ao grupo **Estilos** na faixa de opções do Word enquanto você reduz e expande a janela do Word.) Sobre essa marcação, observe:

    - Ambas as propriedades são necessárias.
    - A `size` unidade de medida da propriedade é pixels. Os ícones são sempre quadrados, portanto, o número é a altura e a largura.
    - A `sourceLocation` propriedade especifica a URL completa para o ícone.

    > [!IMPORTANT]
    > Assim como você normalmente deve alterar as URLs no manifesto do complemento quando você muda do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.

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

1. No nosso exemplo simples em andamento, o grupo tem apenas um único botão. Adicione o objeto a seguir como o único membro da `controls` matriz. Sobre essa marcação, observe:

    - Todas as propriedades, exceto `enabled` , são necessárias.
    - `type` especifica o tipo de controle. Os valores podem ser "Button", "Menu" ou "MobileButton".
    - `id` pode ter até 125 caracteres.
    - `actionId` deve ser a ID de uma ação definida na `actions` matriz. (Consulte a etapa 1 desta seção.)
    - `label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.
    - `superTip` representa uma forma rica de dica de ferramenta. As propriedades `title` e `description` são necessárias.
    - `icon` especifica os ícones do botão. Os comentários anteriores sobre o ícone de grupo também se aplicam aqui.
    - `enabled` (opcional) especifica se o botão está habilitado quando a guia contextual aparece é iniciada. O padrão se não estiver presente é `true` .

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

A seguir está o exemplo completo do blob JSON.

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Registrar a guia contextual com Office requestCreateControls

A guia contextual é registrada com Office chamando o [método Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Isso normalmente é feito na função atribuída ou `Office.initialize` com o `Office.onReady` método. Para saber mais sobre esses métodos e inicializar o add-in, consulte [Initialize your Office Add-in](../develop/initialize-add-in.md). No entanto, você pode chamar o método a qualquer momento após a inicialização.

> [!IMPORTANT]
> O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um complemento. Um erro será lançado se for chamado novamente.

Apresentamos um exemplo a seguir. Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o método antes de poder ser passada para `JSON.parse` uma função JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Especifique os contextos quando a guia ficará visível com requestUpdate

Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto do complemento. Considere um cenário no qual a guia deve estar visível quando, e somente quando, um gráfico (na planilha padrão de uma pasta de trabalho Excel) é ativado.

Comece atribuindo manipuladores. Isso é geralmente feito no método como no exemplo a seguir que atribui manipuladores (criados em uma etapa posterior) aos eventos e de todos os `Office.onReady` `onActivated` `onDeactivated` gráficos na planilha.

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

Em seguida, defina os manipuladores. Veja a seguir um exemplo simples de um , mas consulte Manipulando o `showDataTab` [erro HostRestartNeeded](#handle-the-hostrestartneeded-error) posteriormente neste artigo para obter uma versão mais robusta da função. Sobre este código, observe:

- O Office controla quando atualiza o estado da faixa de opções. O [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestUpdate_input_) enfileia uma solicitação para atualizar. O método resolverá o objeto assim que tiver enraizado a solicitação, não quando a faixa `Promise` de opções realmente for atualizada.
- O parâmetro para o método é um objeto `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia por sua ID exatamente como especificado no *JSON* e (2) especifica a visibilidade da guia.
- Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar objetos de tabulação adicionais à `tabs` matriz.

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

O manipulador para ocultar a guia é quase idêntico, exceto pelo fato de que ela define a `visible` propriedade de volta como `false` .

A Office javaScript também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto. A seguir está `showDataTab` a função em TypeScript e faz uso desses tipos.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Visibilidade da guia de alternância e o status habilitado de um botão ao mesmo tempo

O método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada `requestUpdate` ou em uma guia principal personalizada. Para obter detalhes sobre isso, consulte [Enable and Disable Add-in Commands](disable-add-in-commands.md). Pode haver cenários nos quais você deseja alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo. Você faz isso com uma única chamada de `requestUpdate` . A seguir está um exemplo no qual um botão em uma guia principal é habilitado ao mesmo tempo em que uma guia contextual é visível.

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

No exemplo a seguir, o botão habilitado está na mesma guia contextual que está sendo visível.

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

## <a name="open-a-task-pane-from-contextual-tabs"></a>Abra um painel de tarefas de guias contextuais

Para abrir o painel de tarefas de um botão em uma guia contextual personalizada, crie uma ação no JSON com `type` um `ShowTaskpane` de . Em seguida, defina um botão `actionId` com a propriedade definida como a da `id` ação. Isso abre o painel de tarefas padrão especificado pelo `<Runtime>` elemento em seu manifesto.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

Para abrir qualquer painel de tarefas que não seja o painel de tarefas padrão, especifique uma propriedade na `sourceLocation` definição da ação. No exemplo a seguir, um segundo painel de tarefas é aberto a partir de um botão diferente.

> [!IMPORTANT]
>
> - Quando um é especificado para a ação, o painel de tarefas não usa o tempo de execução `sourceLocation` compartilhado.  Ele é executado em um novo tempo de execução JavaScript.
> - Não mais do que um painel de tarefas pode usar o tempo de execução compartilhado, portanto, mais de uma ação de tipo `ShowTaskpane` pode omitir a `sourceLocation` propriedade.

```json
`{
  "actions": [
    {
      "id": "openChartsTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Charts",
      "supportPinning": false
    },
    {
      "id": "openTablesTaskpane",
      "type": "ShowTaskpane",
      "title": "Work with Tables",
      "supportPinning": false
      "sourceLocation": "https://MyDomain.com/myPage.html"
    }
  ],
  "tabs": [
    {
      // some tab properties omitted
      "groups": [
        {
          // some group properties omitted
          "controls": [
            {
                "type": "Button",
                "id": "CtxBt112",
                "actionId": "openChartsTaskpane",
                "enabled": false,
                "label": "Open Charts Taskpane",
                // some control properties omitted
            },
            {
                "type": "Button",
                "id": "CtxBt113",
                "actionId": "openTablesTaskpane",
                "enabled": false,
                "label": "Open Tables Taskpane",
                // some control properties omitted
            }
          ]
        }
      ]
    }
  ]
}`
```

## <a name="localize-the-json-text"></a>Localize o texto JSON

O blob JSON que é passado não é localizado da mesma maneira que a marcação de manifesto para guias principais personalizadas é localizada (que é descrita em Localização de Controle do `requestCreateControls` [manifesto](../develop/localization.md#control-localization-from-the-manifest)). Em vez disso, a localização deve ocorrer em tempo de execução usando blobs JSON distintos para cada localidade. Sugerimos que você use uma instrução que testa a `switch` [propriedade Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Apresentamos um exemplo a seguir.

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

Em seguida, seu código chama a função para obter o blob localizado que é passado para `requestCreateControls` , como no exemplo a seguir.

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Práticas recomendadas para guias contextuais personalizadas

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implementar uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas

Algumas combinações de plataforma, Office aplicativo e Office build não suportam `requestCreateControls` . Seu complemento deve ser projetado para fornecer uma experiência alternativa aos usuários que estão executando o complemento em uma dessas combinações. As seções a seguir descrevem duas maneiras de fornecer uma experiência de fallback.

#### <a name="use-noncontextual-tabs-or-controls"></a>Usar guias ou controles nãocontextuais

Há um elemento de manifesto, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), projetado para criar uma experiência de fallback em um complemento que implementa guias contextuais personalizadas quando o add-in está sendo executado em um aplicativo ou plataforma que não oferece suporte a guias contextuais personalizadas.

A estratégia mais simples para usar esse elemento é que você define no manifesto uma ou mais guias principais personalizadas (ou seja, guias personalizadas *nãocontextuais)* que duplicam as personalizações da faixa de opções das guias contextuais personalizadas no seu complemento. Mas você adiciona como o primeiro elemento filho dos elementos de grupo, controle e menu duplicados nas `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` guias principais [](../reference/manifest/group.md) [](../reference/manifest/control.md) `<Item>` personalizadas. O efeito de fazer isso é o seguinte:

- Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, os grupos e controles principais personalizados não aparecerão na faixa de opções. Em vez disso, a guia contextual personalizada será criada quando o complemento chamar o `requestCreateControls` método.
- Se o complemento for executado em  um aplicativo ou plataforma que não oferece suporte, os elementos aparecerão nas `requestCreateControls` guias principais personalizadas.

Apresentamos um exemplo a seguir. Observe que "MyButton" aparecerá na guia principal personalizada somente quando as guias contextuais personalizadas não são suportadas. Mas o grupo pai e a guia principal personalizada aparecerão independentemente de as guias contextuais personalizadas são suportadas.

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>              
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Para obter mais exemplos, consulte [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).

Quando um grupo pai ou menu é marcado com , ele não fica visível e toda a marcação filha é ignorada quando as guias contextuais personalizadas não são `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` suportadas. Portanto, não importa se qualquer um desses elementos filho tem o `<OverriddenByRibbonApi>` elemento ou qual é seu valor. A implicação disso é que, se um item de menu ou controle deve estar visível em todos os contextos, então não só não deve ser marcado com , mas seu menu e grupo ancestral também não devem ser marcados `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *dessa forma*.

> [!IMPORTANT]
> Não marque todos *os* elementos filho de um grupo ou menu com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . Isso não faz sentido se o elemento pai for marcado com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` por motivos dados no parágrafo anterior. Além disso, se você deixar de fora o no pai (ou defini-lo como ), o pai aparecerá independentemente de as guias contextuais personalizadas serem suportadas, mas elas estarão vazias quando elas são `<OverriddenByRibbonApi>` `false` suportadas. Portanto, se todos os elementos filho não aparecerem quando as guias contextuais personalizadas são suportadas, marque o pai e somente o pai, com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Usar APIs que mostram ou ocultam um painel de tarefas em contextos especificados

Como alternativa a , seu complemento pode definir um painel de tarefas com controles de interface do usuário que duplicam a funcionalidade dos controles em uma `<OverriddenByRibbonApi>` guia contextual personalizada. Em seguida, use os métodos [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) [e Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) para mostrar o painel de tarefas quando e somente quando a guia contextual teria sido mostrada se tivesse suporte. Para obter detalhes sobre como usar esses métodos, consulte [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Manipular o erro HostRestartNeeded

Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro. Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto. Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`. Seu código deve lidar com esse erro. Veja a seguir um exemplo de como. Nesse caso, o método `reportError` exibe o erro para o usuário.

```javascript
function showDataTab() {
    try {
        Office.ribbon.requestUpdate({
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
