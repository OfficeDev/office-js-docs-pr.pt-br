---
title: Crie guias contextuais personalizadas em Office Add-ins
description: Aprenda a adicionar guias contextuais personalizadas ao seu Office Add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d03ac2c01c03353f3e2d1b54ba20616d7b42d93f
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555203"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a>Crie guias contextuais personalizadas em Office Add-ins

Uma guia contextual é um controle de guia oculto na fita Office que é exibida na linha de guia quando um evento especificado ocorre no documento Office. Por exemplo, a guia **Design de tabela** que aparece na fita Excel quando uma tabela é selecionada. Você pode incluir guias contextuais personalizadas em seu Office Add-in e especificar quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade. (No entanto, as guias contextuais personalizadas não respondem às alterações de foco.)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

> [!IMPORTANT]
> Atualmente, as guias contextuais personalizadas são suportadas apenas em Excel e somente nessas plataformas e compilações:
>
> - Excel em Windows (somente Microsoft 365 assinatura): Versão 2102 (Build 13801.20294) ou posterior.
> - Excel Online

> [!NOTE]
> As guias contextuais personalizadas funcionam apenas em plataformas que suportam os seguintes conjuntos de requisitos. Para obter mais informações sobre os conjuntos de requisitos e como trabalhar com eles, consulte [Especificar Office aplicativos e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [RibbonApi 1.2](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [SharedRuntime 1.1](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> Você pode usar as verificações de tempo de execução em seu código para testar se a combinação de host e plataforma do usuário suporta esses conjuntos de requisitos conforme descrito em [Especificar Office aplicativos e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). (A técnica de especificar os conjuntos de exigências no manifesto, que também está descrito nesse artigo, não funciona atualmente para RibbonApi 1.2.) Alternativamente, você pode [implementar uma experiência de interface do usuário alternativa quando as guias contextuais personalizadas não forem suportadas](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

## <a name="behavior-of-custom-contextual-tabs"></a>Comportamento de guias contextuais personalizadas

A experiência do usuário para guias contextuais personalizadas segue o padrão de guias Office contextuais incorporadas. A seguir, os princípios básicos para as guias contextuais personalizadas de colocação:

- Quando uma guia contextual personalizada é visível, ela aparece na extremidade direita da fita.
- Se uma ou mais guias contextuais incorporadas e uma ou mais guias contextuais personalizadas de complementos forem visíveis ao mesmo tempo, as guias contextuais personalizadas estão sempre à direita de todas as guias contextuais incorporadas.
- Se o seu complemento tiver mais de uma guia contextual e houver contextos em que mais de um seja visível, eles aparecem na ordem em que são definidos em seu complemento. (A direção é a mesma direção que a língua Office; ou seja, é da esquerda para a direita em línguas da esquerda para a direita, mas da direita para a esquerda em línguas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como você os define.
- Se mais de um complemento tiver uma guia contextual visível em um contexto específico, então eles aparecem na ordem em que os complementos foram lançados.
- As guias *contextuais* personalizadas, ao contrário das guias de núcleo personalizadas, não são adicionadas permanentemente à fita do aplicativo Office. Eles estão presentes apenas em Office documentos sobre os quais seu complemento está sendo executado.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Principais etapas para incluir uma guia contextual em um complemento

A seguir, os principais passos para incluir uma guia contextual personalizada em um complemento:

1. Configure o complemento para usar um tempo de execução compartilhado.
1. Defina a guia e os grupos e controles que aparecem nela.
1. Cadastre-se na aba contextual com Office.
1. Especifique as circunstâncias em que a guia será visível.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configure o complemento para usar um tempo de execução compartilhado

Adicionar guias contextuais personalizadas requer que seu complemento use o tempo de execução compartilhado. Para obter mais informações, consulte [Configurar um complemento para usar um tempo de execução compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Defina os grupos e controles que aparecem na guia

Ao contrário das guias de núcleo personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com uma bolha JSON. Seu código analisa a bolha em um objeto JavaScript e, em seguida, passa o objeto para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) As guias contextuais personalizadas só estão presentes em documentos nos quais seu complemento está sendo executado no momento. Isso é diferente das guias de núcleo personalizadas que são adicionadas à fita de aplicação Office quando o complemento é instalado e permanecem presentes quando outro documento é aberto. Além disso, o `requestCreateControls` método pode ser executado apenas uma vez em uma sessão do seu complemento. Se for chamado novamente, um erro é jogado.

> [!NOTE]
> A estrutura das propriedades e subpropriedades do blob JSON (e os nomes-chave) é aproximadamente paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no manifesto XML.

Construiremos um exemplo de uma aba contextual JSON blob passo a passo. O esquema completo para a aba contextual JSON está em [dynamic-ribbon.schema.jsem.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json) Se você estiver trabalhando em Visual Studio Code, você pode usar este arquivo para obter IntelliSense e validar seu JSON. Para obter mais informações, consulte [Editando JSON com Visual Studio Code - esquemas e configurações JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Comece criando uma sequência JSON com duas propriedades de matriz `actions` nomeadas e `tabs` . O `actions` array é uma especificação de todas as funções que podem ser executadas por controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, até um máximo de *20*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Este simples exemplo de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação. Adicione o seguinte como o único membro da `actions` matriz. Sobre esta marcação, nota:

    - As `id` `type` propriedades são obrigatórias.
    - O valor `type` pode ser "ExecuteFunction" ou "ShowTaskpane".
    - O `functionName` imóvel só é usado quando o valor é `type` `ExecuteFunction` . É o nome de uma função definida no FunctionFile. Para obter mais informações sobre o FunctionFile, consulte [conceitos básicos para comandos adicionais](add-in-commands.md).
    - Em um passo posterior, você mapeará esta ação para um botão na guia contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Adicione o seguinte como o único membro da `tabs` matriz. Sobre esta marcação, nota:

    - A propriedade `id` é obrigatória. Use um ID breve e descritivo que seja único entre todas as guias contextuais em seu complemento.
    - A propriedade `label` é obrigatória. É uma sequência fácil de usar para servir como o rótulo da guia contextual.
    - A propriedade `groups` é obrigatória. Ele define os grupos de controles que aparecerão na guia. Deve ter pelo menos um membro *e não mais de 20.* (Há também limites no número de controles que você pode ter em uma guia contextual personalizada e isso também restringirá quantos grupos você tem. Veja o próximo passo para obter mais informações.)

    > [!NOTE]
    > O objeto da guia também pode ter uma propriedade opcional `visible` que especifica se a guia é visível imediatamente quando o complemento é iniciado. Uma vez que as guias contextuais são normalmente ocultas até que um evento de usuário acione sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), o `visible` imóvel padrão para quando não está `false` presente. Em uma seção posterior, mostramos como definir a propriedade `true` em resposta a um evento.

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. No simples exemplo em curso, a guia contextual tem apenas um único grupo. Adicione o seguinte como o único membro da `groups` matriz. Sobre esta marcação, nota:

    - Todas as propriedades são necessárias.
    - A `id` propriedade deve ser única entre todos os grupos da guia. Use um Y breve e descritivo.
    - A `label` é uma string fácil de usar para servir como o rótulo do grupo.
    - O `icon` valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na fita, dependendo do tamanho da fita e da janela de aplicação Office.
    - O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e menus do grupo. Deve haver pelo menos um.

    > [!IMPORTANT]
    > *O número total de controles na guia geral não pode ser superior a 20.* Por exemplo, você poderia ter 3 grupos com 6 controles cada, e um quarto grupo com 2 controles, mas você não pode ter 4 grupos com 6 controles cada.  

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

1. Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32x32 px e 80x80 px. Opcionalmente, você também pode ter ícones dos tamanhos 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px. Office decide qual ícone usar com base no tamanho da fita e Office janela de aplicativo. Adicione os seguintes objetos à matriz de ícones. (Se os tamanhos da janela e da fita forem grandes o suficiente para que pelo menos um dos *controles* do grupo apareça, então nenhum ícone de grupo aparece. Por exemplo, assista ao grupo **Styles** na fita do Word enquanto você encolhe e expande a janela do Word.) Sobre esta marcação, nota:

    - Ambas as propriedades são necessárias.
    - A `size` unidade de propriedade da medida é pixels. Os ícones são sempre quadrados, então o número é tanto a altura quanto a largura.
    - A `sourceLocation` propriedade especifica a URL completa para o ícone.

    > [!IMPORTANT]
    > Assim como você normalmente deve alterar os URLs no manifesto do complemento quando você passar do desenvolvimento para a produção (como mudar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.

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

1. Em nosso simples exemplo contínuo, o grupo tem apenas um único botão. Adicione o seguinte objeto como o único membro da `controls` matriz. Sobre esta marcação, nota:

    - Todas as propriedades, `enabled` exceto, são necessárias.
    - `type` especifica o tipo de controle. Os valores podem ser "Button", "Menu" ou "MobileButton".
    - `id` pode ter até 125 caracteres. 
    - `actionId` deve ser o ID de uma ação definida na `actions` matriz. (Veja o passo 1 desta seção.)
    - `label` é uma string fácil de usar para servir como a etiqueta do botão.
    - `superTip` representa uma rica forma de ponta de ferramenta. Tanto as propriedades quanto as `title` `description` propriedades são necessárias.
    - `icon` especifica os ícones para o botão. As observações anteriores sobre o ícone de grupo também se aplicam aqui.
    - `enabled` (opcional) especifica se o botão está ativado quando a guia contextual é ativada. O padrão se não estiver presente é `true` . 

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
 
O seguinte é o exemplo completo da bolha JSON:

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Cadastre-se na guia contextual com Office solicitaçãoCreateControls

A guia contextual é registrada com Office ligando para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) Isso é normalmente feito na função que é atribuída `Office.initialize` ou com o `Office.onReady` método. Para obter mais informações sobre esses métodos e inicializar o complemento, consulte [Initialize seu Office Add-in](../develop/initialize-add-in.md). Você pode, no entanto, chamar o método a qualquer momento após a inicialização.

> [!IMPORTANT]
> O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um complemento. Um erro é jogado se for chamado novamente.

Apresentamos um exemplo a seguir. Observe que a sequência JSON deve ser convertida em um objeto JavaScript com o `JSON.parse` método antes que possa ser passado para uma função JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Especifique os contextos quando a guia estará visível com a solicitaçãoUpdate

Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto de complementação. Considere um cenário em que a guia deve ser visível quando, e somente quando, um gráfico (na planilha padrão de uma Excel pasta de trabalho) é ativado.

Comece designando manipuladores. Isso é comumente feito no `Office.onReady` método como no exemplo a seguir que atribui manipuladores (criados em uma etapa posterior) aos eventos `onActivated` de todos os `onDeactivated` gráficos na planilha.

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

Em seguida, defina os manipuladores. A seguir, um exemplo simples de um `showDataTab` , mas veja [Manipulação do erro HostRestartNeed](#handle-the-hostrestartneeded-error) mais tarde neste artigo para uma versão mais robusta da função. Sobre este código, observe:

- O Office controla quando atualiza o estado da faixa de opções. O método [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) faz filas em uma solicitação para atualizar. O método resolverá o `Promise` objeto assim que ele tiver enfileido a solicitação, não quando a fita realmente se atualizar.
- O parâmetro para o `requestUpdate` método é um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia pelo seu ID *exatamente conforme especificado no JSON* e (2) especifica a visibilidade da guia.
- Se você tiver mais de uma guia contextual personalizada que deve ser visível no mesmo contexto, basta adicionar objetos de guia adicionais à `tabs` matriz.

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

O manipulador para ocultar a guia é quase idêntico, exceto que ele define a `visible` propriedade de volta para `false` .

A biblioteca Office JavaScript também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto. A seguir, a `showDataTab` função no TypeScript e faz uso desses tipos.

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Alternar a visibilidade da guia e o status habilitado de um botão ao mesmo tempo

O `requestUpdate` método também é usado para alternar o status ativado ou desativado de um botão personalizado em uma guia contextual personalizada ou em uma guia central personalizada. Para obter detalhes sobre isso, consulte [Ativar e Desativar comandos adicionais](disable-add-in-commands.md). Pode haver cenários em que você deseja alterar tanto a visibilidade de uma guia quanto o status habilitado de um botão ao mesmo tempo. Você pode fazer isso com uma única chamada de `requestUpdate` . A seguir, um exemplo no qual um botão em uma guia central é ativado ao mesmo tempo que uma guia contextual é visível.

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

No exemplo a seguir, o botão habilitado está na mesma aba contextual que está sendo visível.

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

## <a name="localizing-the-json-blob"></a>Localização do blob JSON

A bolha JSON que é passada `requestCreateControls` não é localizada da mesma forma que a marcação manifesto para guias de núcleo personalizadas é localizada (que é descrita na localização do Controle a partir do [manifesto](../develop/localization.md#control-localization-from-the-manifest)). Em vez disso, a localização deve ocorrer no tempo de execução usando bolhas JSON distintas para cada localidade. Sugerimos que você use uma `switch` instrução que testa a propriedade [Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage) Veja um exemplo a seguir:

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

Em seguida, seu código chama a função para obter a bolha localizada que é passada para `requestCreateControls` , como no exemplo a seguir:

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a>Melhores práticas para guias contextuais personalizadas

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a>Implemente uma experiência de interface do usuário alternativa quando as guias contextuais personalizadas não forem suportadas

Algumas combinações de plataforma, Office aplicativo e Office build não suportam `requestCreateControls` . Seu complemento deve ser projetado para fornecer uma experiência alternativa aos usuários que estão executando o complemento em uma dessas combinações. As seções a seguir descrevem duas maneiras de fornecer uma experiência de recuo.

#### <a name="use-noncontextual-tabs-or-controls"></a>Use guias ou controles não contextuais

Há um elemento manifesto, [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)que foi projetado para criar uma experiência de recuo em um complemento que implementa guias contextuais personalizadas quando o complemento está sendo executado em um aplicativo ou plataforma que não suporta guias contextuais personalizadas. 

A estratégia mais simples para usar este elemento é que você define nas guias de núcleo manifesto ou mais personalizadas (ou seja, guias personalizadas *não contextuais)* que duplicam as personalizações de fita das guias contextuais personalizadas em seu complemento. Mas você adiciona `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` como o elemento primeiro filho do [CustomTab](../reference/manifest/customtab.md). O efeito de fazê-lo é o seguinte:

- Se o complemento for executado em um aplicativo e plataforma que suportem guias contextuais personalizadas, a guia central personalizada não aparecerá na fita. Em vez disso, a guia contextual personalizada será criada quando o complemento chamar o `requestCreateControls` método.
- Se o complemento for executado em um aplicativo ou plataforma que *não* `requestCreateControls` suporte, a guia central personalizada aparecerá na fita.

A seguir, um exemplo dessa simples estratégia.

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
              <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
...
</OfficeApp>
```

Esta estratégia simples usa uma guia central personalizada que espelha uma guia contextual personalizada com seus grupos e controles infantis, mas você pode usar uma estratégia mais complexa. O `<OverriddenByRibbonApi>` elemento também pode ser adicionado como (o primeiro) elemento infantil aos elementos [grupo](../reference/manifest/group.md) e [controle](../reference/manifest/control.md) (tanto tipo [de botão](../reference/manifest/control.md#button-control) quanto tipo de [menu](../reference/manifest/control.md#menu-dropdown-button-controls)), e elementos do `<Item>` menu. Este fato permite que você distribua os grupos e controles que de outra forma apareceriam na guia contextual entre vários grupos, botões e menus em várias guias de núcleo personalizadas. Apresentamos um exemplo a seguir. Observe que "MyButton" aparecerá na guia principal personalizada somente quando as guias contextuais personalizadas não forem suportadas. Mas o grupo pai e a guia central personalizada aparecerão independentemente de as guias contextuais personalizadas serem suportadas.

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

Quando uma guia, grupo ou menu dos pais é marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` com , então ele não é visível, e toda a marcação de criança é ignorada, quando as guias contextuais personalizadas não são suportadas. Então, não importa se algum desses elementos infantis tem o `<OverriddenByRibbonApi>` elemento ou qual é o seu valor. A implicação disso é que se um item, controle ou grupo do menu deve ser visível em todos os contextos, então não só não deve ser marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` com , mas seu *menu, grupo e guia ancestrais também não devem ser marcados dessa forma*.

> [!IMPORTANT]
> Não marque *todos os* elementos infantis de uma guia, grupo ou menu com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` . Isso é inútil se o elemento pai estiver marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` por razões dadas no parágrafo anterior. Além disso, se você deixar de fora `<OverriddenByRibbonApi>` o pai (ou defini-lo `false` para ), então o pai aparecerá independentemente de as guias contextuais personalizadas serem suportadas, mas estará vazia quando elas forem suportadas. Assim, se todos os elementos da criança não aparecerem quando as guias contextuais personalizadas forem suportadas, marque o pai e apenas o pai, com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a>Use APIs que mostram ou ocultam um painel de tarefas em contextos especificados

Como alternativa, `<OverriddenByRibbonApi>` seu complemento pode definir um painel de tarefas com controles de interface do usuário que duplicam a funcionalidade dos controles em uma guia contextual personalizada. Em seguida, use os métodos [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) e [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) para mostrar o painel de tarefas quando, e somente quando, a guia contextual teria sido mostrada se fosse suportada. Para obter detalhes sobre como usar esses métodos, consulte [Mostrar ou ocultar o painel de tarefas do seu Office Add-in](../develop/show-hide-add-in.md).

### <a name="handle-the-hostrestartneeded-error"></a>Manuseie o erro HostRestartNeed

Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro. Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto. Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`. Seu código deve lidar com esse erro. A seguir, um exemplo de como. Nesse caso, o método `reportError` exibe o erro para o usuário.

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
