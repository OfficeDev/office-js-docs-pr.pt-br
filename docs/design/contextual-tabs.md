---
title: Criar guias contextuais personalizadas em Suplementos do Office
description: Saiba como adicionar guias contextuais personalizadas ao suplemento do Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f43f6ec0a6ef3faef4c5e50d5da6d124124fe92
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810229"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Criar guias contextuais personalizadas em Suplementos do Office

Uma guia contextual é um controle de guia oculto na faixa de opções do Office exibida na linha de guias quando ocorre um evento especificado no documento do Office. Por exemplo, a guia **Design de Tabela** que aparece na faixa de opções do Excel quando uma tabela é selecionada. Você inclui guias contextuais personalizadas em seu Suplemento do Office e especifica quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade. (No entanto, as guias contextuais personalizadas não respondem às alterações de foco.)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

> [!IMPORTANT]
> Atualmente, as guias contextuais personalizadas só têm suporte no Excel e somente nessas plataformas e builds.
>
> - Excel no Windows: versão 2102 (Build 13801.20294) ou posterior.
> - Excel no Mac: versão 16.53.806.0 ou posterior.
> - Excel Online

> [!NOTE]
> As guias contextuais personalizadas funcionam apenas em plataformas que dão suporte aos seguintes conjuntos de requisitos. Para saber mais sobre os conjuntos de requisitos e como trabalhar com eles, consulte [Especificar aplicativos do Office e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
> - [SharedRuntime 1.1](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)
>
> Você pode usar as verificações de runtime em seu código para testar se a combinação de host e plataforma do usuário dá suporte a esses conjuntos de requisitos, conforme descrito nas [verificações do Runtime para o método e o suporte ao conjunto de requisitos](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (A técnica de especificar os conjuntos de requisitos no manifesto, que também é descrito nesse artigo, não funciona atualmente para RibbonApi 1.2.) Como alternativa, você pode [implementar uma experiência alternativa de interface do usuário quando não há suporte para guias contextuais personalizadas](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportamento de guias contextuais personalizadas

A experiência do usuário para guias contextuais personalizadas segue o padrão de guias contextuais internas do Office. A seguir estão os princípios básicos para as guias contextuais personalizadas de posicionamento.

- Quando uma guia contextual personalizada é visível, ela é exibida na extremidade direita da faixa de opções.
- Se uma ou mais guias contextuais internas e uma ou mais guias contextuais personalizadas de suplementos estiverem visíveis ao mesmo tempo, as guias contextuais personalizadas estarão sempre à direita de todas as guias contextuais internas.
- Se o suplemento tiver mais de uma guia contextual e houver contextos em que mais de um esteja visível, eles aparecerão na ordem em que são definidos no suplemento. (A direção é a mesma direção que a linguagem do Office; ou seja, é da esquerda para a direita em idiomas da esquerda para a direita, mas da direita para a esquerda em idiomas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como defini-los.
- Se mais de um suplemento tiver uma guia contextual visível em um contexto específico, eles aparecerão na ordem em que os suplementos foram iniciados.
- As guias *contextuais* personalizadas, ao contrário das guias principais personalizadas, não são adicionadas permanentemente à faixa de opções do aplicativo do Office. Eles estão presentes apenas em documentos do Office nos quais seu suplemento está em execução.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Principais etapas para incluir uma guia contextual em um suplemento

A seguir estão as principais etapas para incluir uma guia contextual personalizada em um suplemento.

1. Configure o suplemento para usar um runtime compartilhado.
1. Defina a guia e os grupos e controles que aparecem nela.
1. Registre a guia contextual com o Office.
1. Especifique as circunstâncias em que a guia estará visível.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o suplemento para usar um runtime compartilhado

Adicionar guias contextuais personalizadas requer que seu suplemento use o [runtime compartilhado](../testing/runtimes.md#shared-runtime). Para obter mais informações, consulte [Configurar um suplemento para usar um runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Definir os grupos e controles que aparecem na guia

Ao contrário das guias principais personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas no runtime com um blob JSON. Seu código analisa o blob em um objeto JavaScript e, em seguida, passa o objeto para o método [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . As guias contextuais personalizadas só estão presentes em documentos nos quais o suplemento está em execução no momento. Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções de aplicativo do Office quando o suplemento é instalado e permanecem presentes quando outro documento é aberto. Além disso, o `requestCreateControls` método pode ser executado apenas uma vez em uma sessão do seu suplemento. Se ele for chamado novamente, um erro será gerado.

> [!NOTE]
> A estrutura das propriedades e subpropertidades do blob JSON (e os nomes de chave) é aproximadamente paralela à estrutura do elemento [CustomTab](/javascript/api/manifest/customtab) e seus elementos descendentes no XML de manifesto.

Construiremos um exemplo de um blob JSON de guias contextuais passo a passo. O esquema completo da guia contextual JSON está em [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Se você estiver trabalhando em Visual Studio Code, poderá usar esse arquivo para obter o IntelliSense e validar seu JSON. Para obter mais informações, consulte [Edição de JSON com Visual Studio Code – esquemas JSON e configurações](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).

1. Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz nomeadas `actions` e `tabs`. A `actions` matriz é uma especificação de todas as funções que podem ser executadas por controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, *até um máximo de 20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação. Adicione o seguinte como o único membro da `actions` matriz. Sobre essa marcação, observe:

    - As `id` propriedades e `type` são obrigatórias.
    - O valor de `type` pode ser "ExecuteFunction" ou "ShowTaskpane".
    - A `functionName` propriedade só é usada quando o valor de `type` é `ExecuteFunction`. É o nome de uma função definida no FunctionFile. Para obter mais informações sobre o FunctionFile, consulte [Conceitos básicos para comandos de suplemento](add-in-commands.md).
    - Em uma etapa posterior, você mapeará essa ação para um botão na guia contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
    ```

1. Adicione o seguinte como o único membro da `tabs` matriz. Sobre essa marcação, observe:

    - A propriedade `id` é obrigatória. Use uma ID breve e descritiva exclusiva entre todas as guias contextuais no suplemento.
    - A propriedade `label` é obrigatória. É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.
    - A propriedade `groups` é obrigatória. Ele define os grupos de controles que serão exibidos na guia. Ele deve ter pelo menos um membro *e não mais de 20*. (Também há limites no número de controles que você pode ter em uma guia contextual personalizada e isso também restringirá quantos grupos você tem. Confira a próxima etapa para obter mais informações.)

    > [!NOTE]
    > O objeto tab também pode ter uma propriedade opcional `visible` que especifica se a guia fica visível imediatamente quando o suplemento é iniciado. Como as guias contextuais normalmente são ocultas até que um evento de usuário dispare sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a `visible` propriedade é padrão quando `false` não está presente. Em uma seção posterior, mostramos como definir a propriedade como `true` em resposta a um evento.

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
    - A `id` propriedade deve ser exclusiva entre todos os grupos no manifesto. Use uma ID breve e descritiva de até 125 caracteres.
    - A `label` é uma cadeia de caracteres amigável para servir como o rótulo do grupo.
    - O `icon` valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela do aplicativo do Office.
    - O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e menus no grupo. Deve haver pelo menos um.

    > [!IMPORTANT]
    > *O número total de controles na guia inteira não pode ser superior a 20.* Por exemplo, você pode ter três grupos com 6 controles cada e um quarto grupo com dois controles, mas não pode ter quatro grupos com 6 controles cada.  

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

1. Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32x32 px e 80x80 px. Opcionalmente, você também pode ter ícones de tamanhos 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px. O Office decide qual ícone usar com base no tamanho da faixa de opções e da janela do aplicativo do Office. Adicione os objetos a seguir à matriz de ícones. (Se os tamanhos da janela e da faixa de opções forem grandes o suficiente para que pelo menos um dos *controles* do grupo apareça, nenhum ícone de grupo será exibido. Para obter um exemplo, assista ao grupo **Styles** na faixa de opções do Word enquanto você reduz e expande a janela do Word.) Sobre essa marcação, observe:

    - Ambas as propriedades são necessárias.
    - A `size` unidade de propriedade da medida é pixels. Os ícones são sempre quadrados, portanto, o número é a altura e a largura.
    - A `sourceLocation` propriedade especifica a URL completa para o ícone.

    > [!IMPORTANT]
    > Assim como normalmente você deve alterar as URLs no manifesto do suplemento ao passar do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.

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

1. Em nosso exemplo simples em andamento, o grupo tem apenas um único botão. Adicione o objeto a seguir como o único membro da `controls` matriz. Sobre essa marcação, observe:

    - Todas as propriedades, exceto `enabled`, são necessárias.
    - `type` especifica o tipo de controle. Os valores podem ser "Button", "Menu" ou "MobileButton".
    - `id` pode ter até 125 caracteres.
    - `actionId` deve ser a ID de uma ação definida na `actions` matriz. (Confira a etapa 1 desta seção.)
    - `label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.
    - `superTip` representa uma forma avançada de dica de ferramenta. `title` As propriedades e `description` são necessárias.
    - `icon` especifica os ícones do botão. As observações anteriores sobre o ícone de grupo também se aplicam aqui.
    - `enabled` (opcional) especifica se o botão está habilitado quando a guia contextual é exibida é iniciada. O padrão se não estiver presente é `true`.

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Registrar a guia contextual com o Office com requestCreateControls

A guia contextual é registrada no Office chamando o método [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestcreatecontrols-member(1)) . Normalmente, isso é feito na função atribuída a `Office.initialize` ou com a `Office.onReady` função. Para saber mais sobre essas funções e inicializar o suplemento, consulte [Inicializar seu Suplemento do Office](../develop/initialize-add-in.md). No entanto, você pode chamar o método a qualquer momento após a inicialização.

> [!IMPORTANT]
> O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um suplemento. Um erro será gerado se ele for chamado novamente.

Apresentamos um exemplo a seguir. Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o `JSON.parse` método antes de ser passada para uma função JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Especifique os contextos quando a guia ficará visível com requestUpdate

Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto de suplemento. Considere um cenário no qual a guia deve estar visível quando e somente quando um gráfico (na planilha padrão de uma pasta de trabalho do Excel) é ativado.

Comece atribuindo manipuladores. Isso geralmente é feito na `Office.onReady` função como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos `onActivated` eventos e `onDeactivated` de todos os gráficos na planilha.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);

    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(showDataTab);
        charts.onDeactivated.add(hideDataTab);
        return context.sync();
    });
});
```

Em seguida, defina os manipuladores. A seguir está um exemplo simples de um `showDataTab`, mas consulte [Manipulando o erro HostRestartNeeded](#handle-the-hostrestartneeded-error) mais tarde neste artigo para obter uma versão mais robusta da função. Sobre este código, observe:

- O Office controla quando atualiza o estado da faixa de opções. O método  [Office.ribbon.requestUpdate enfileira](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)) uma solicitação para atualizar. O método resolverá o `Promise` objeto assim que ele tiver enfileirado a solicitação, não quando a faixa de opções realmente for atualizada.
- O parâmetro para o `requestUpdate` método é um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia por sua ID *exatamente conforme especificado no JSON* e (2) especifica a visibilidade da guia.
- Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar objetos de guia adicionais à `tabs` matriz.

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

O manipulador para ocultar a guia é quase idêntico, exceto que ele define a `visible` propriedade de volta como `false`.

A biblioteca JavaScript do Office também fornece várias interfaces (tipos) para facilitar a construção do`RibbonUpdateData` objeto. A seguir está a `showDataTab` função em TypeScript e faz uso desses tipos.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Visibilidade da guia de alternância e o status habilitado de um botão ao mesmo tempo

O `requestUpdate` método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada ou em uma guia principal personalizada. Para obter detalhes sobre isso, consulte [Habilitar e desabilitar comandos de suplemento](disable-add-in-commands.md). Pode haver cenários em que você deseja alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo. Você faz isso com uma única chamada de `requestUpdate`. Veja a seguir um exemplo em que um botão em uma guia principal é habilitado ao mesmo tempo em que uma guia contextual fica visível.

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

## <a name="open-a-task-pane-from-contextual-tabs"></a>Abrir um painel de tarefas de guias contextuais

Para abrir o painel de tarefas de um botão em uma guia contextual personalizada, crie uma ação no JSON com um `type` de `ShowTaskpane`. Em seguida, defina um botão com a `actionId` propriedade definida como a `id` da ação. Isso abre o painel de tarefas padrão especificado pelo **\<Runtime\>** elemento em seu manifesto.

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

Para abrir qualquer painel de tarefas que não seja o painel de tarefas padrão, especifique uma `sourceLocation` propriedade na definição da ação. No exemplo a seguir, um segundo painel de tarefas é aberto de um botão diferente.

> [!IMPORTANT]
>
> - Quando um `sourceLocation` é especificado para a ação, o painel de tarefas *não* usa o runtime compartilhado. Ele é executado em um novo runtime separado.
> - Não mais do que um painel de tarefas pode usar o runtime compartilhado, portanto, não mais de uma ação do tipo `ShowTaskpane` pode omitir a `sourceLocation` propriedade.

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

## <a name="localize-the-json-text"></a>Localizar o texto JSON

O blob JSON para `requestCreateControls` o qual é passado não é localizado da mesma forma que a marcação de manifesto para guias principais personalizadas é localizada (o que é descrito na [localização de controle do manifesto](../develop/localization.md#control-localization-from-the-manifest)). Em vez disso, a localização deve ocorrer no runtime usando blobs JSON distintos para cada localidade. Sugerimos que você use uma `switch` instrução que teste a propriedade [Office.context.displayLanguage](/javascript/api/office/office.context#office-office-context-displaylanguage-member) . Apresentamos um exemplo a seguir.

```javascript
function GetContextualTabsJsonSupportedLocale () {
    const displayLanguage = Office.context.displayLanguage;

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

Em seguida, seu código chama a função para obter o blob localizado que é passado para `requestCreateControls`, como no exemplo a seguir.

```javascript
const contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Práticas recomendadas para guias contextuais personalizadas

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implementar uma experiência alternativa de interface do usuário quando não há suporte para guias contextuais personalizadas

Algumas combinações de plataforma, aplicativo do Office e build do Office não dão suporte `requestCreateControls`a . Seu suplemento deve ser projetado para fornecer uma experiência alternativa aos usuários que estão executando o suplemento em uma dessas combinações. As seções a seguir descrevem duas maneiras de fornecer uma experiência de fallback.

#### <a name="use-noncontextual-tabs-or-controls"></a>Usar guias ou controles não contratuais

Há um elemento de manifesto, [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi), que foi projetado para criar uma experiência de fallback em um suplemento que implementa guias contextuais personalizadas quando o suplemento está sendo executado em um aplicativo ou plataforma que não dá suporte a guias contextuais personalizadas.

A estratégia mais simples para usar esse elemento é definir uma guia de núcleo personalizada (ou seja, guia personalizada *não contratual* ) no manifesto que duplica as personalizações de faixa de opções das guias contextuais personalizadas no suplemento. Mas você adiciona `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` como o primeiro elemento filho dos elementos duplicados [grupo](/javascript/api/manifest/group), [controle](/javascript/api/manifest/control) e menu **\<Item\>** nas guias principais personalizadas. O efeito de fazer isso é o seguinte:

- Se o suplemento for executado em um aplicativo e plataforma com suporte a guias contextuais personalizadas, os grupos e controles principais personalizados não aparecerão na faixa de opções. Em vez disso, a guia contextual personalizada será criada quando o suplemento chamar o `requestCreateControls` método.
- Se o suplemento for executado em um aplicativo ou plataforma que *não dá* suporte `requestCreateControls`, os elementos aparecerão na guia núcleo personalizado.

Apresentamos um exemplo a seguir. Observe que "MyButton" será exibido na guia núcleo personalizado somente quando não houver suporte para guias contextuais personalizadas. Mas o grupo pai e a guia núcleo personalizado serão exibidos independentemente de as guias contextuais personalizadas serem compatíveis.

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
                <Control ... id="Contoso.MyButton1">
                  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
                  ...
                  <Action ...>
...
</OfficeApp>
```

Para obter mais exemplos, consulte [OverriddenByRibbonApi](/javascript/api/manifest/overriddenbyribbonapi).

Quando um grupo pai ou menu é marcado com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, então ele não é visível e toda a marcação filho é ignorada quando as guias contextuais personalizadas não têm suporte. Portanto, não importa se algum desses elementos filho tem o **\<OverriddenByRibbonApi\>** elemento ou qual é o seu valor. A implicação disso é que, se um item ou controle de menu deve estar visível em todos os contextos, ele não só não deve ser marcado com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, mas *seu menu ancestral e grupo também não devem ser marcados dessa forma*.

> [!IMPORTANT]
> Não marque *todos os* elementos filho de um grupo ou menu com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`. Isso será inútil se o elemento pai for marcado com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` por razões dadas no parágrafo anterior. Além disso, se você deixar de fora o **\<OverriddenByRibbonApi\>** pai (ou defini-lo como `false`), o pai aparecerá independentemente de as guias contextuais personalizadas terem suporte, mas elas ficarão vazias quando tiverem suporte. Portanto, se todos os elementos filho não devem aparecer quando as guias contextuais personalizadas tiverem suporte, marque o pai com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Usar APIs que mostram ou ocultam um painel de tarefas em contextos especificados

Como alternativa ao **\<OverriddenByRibbonApi\>**, o suplemento pode definir um painel de tarefas com controles de interface do usuário que duplicam a funcionalidade dos controles em uma guia contextual personalizada. Em seguida, use os métodos [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-showastaskpane-member(1)) e [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#office-office-addin-hide-member(1)) para mostrar o painel de tarefas quando a guia contextual teria sido mostrada se tivesse suporte. Para obter detalhes sobre como usar esses métodos, consulte [Mostrar ou ocultar o painel de tarefas do suplemento do Office](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Manipular o erro HostRestartNeeded

Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro. Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto. Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`. Seu código deve lidar com esse erro. A seguir está um exemplo de como. Nesse caso, o método `reportError` exibe o erro para o usuário.

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

## <a name="resources"></a>Recursos

- [Exemplo de código: criar guias contextuais personalizadas na faixa de opções](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs)
- Demonstração da comunidade de exemplo de guias contextuais

> [!VIDEO https://www.youtube.com/embed/9tLfm4boQIo]
