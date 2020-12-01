---
title: Criar guias contextuais personalizadas em suplementos do Office
description: Saiba como adicionar guias contextuais personalizadas ao suplemento do Office.
ms.date: 11/20/2020
localization_priority: Normal
ms.openlocfilehash: 49a773aca0651b88c972c24a4cde0aa1e300d5e7
ms.sourcegitcommit: 6619e07cdfa68f9fa985febd5f03caf7aee57d5e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/30/2020
ms.locfileid: "49505550"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a>Criar guias contextuais personalizadas em suplementos do Office (visualização)

Uma guia contextual é um controle de guia oculto na faixa de opções do Office que é exibido na linha da guia quando um evento especificado ocorre no documento do Office. Por exemplo, a guia **design da tabela** que aparece na faixa de opções do Excel quando uma tabela é selecionada. Você pode incluir guias contextuais personalizadas no suplemento do Office e especificar quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade. (No entanto, as guias contextuais personalizadas não respondem às alterações de foco.)

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

> [!IMPORTANT]
> Guias contextuais personalizadas estão em versão prévia. Faça experiências com eles em um ambiente de desenvolvimento ou teste, mas não os adicione a um suplemento de produção.
>
> Atualmente, as guias contextuais personalizadas só têm suporte no Excel e apenas nessas plataformas e compilações:
>
> - Excel no Windows (somente Microsoft 365, licença não permanente): versão 2011 (Build 13426,20274). Sua assinatura do Microsoft 365 pode precisar estar no [canal atual (visualização)](https://insider.office.com/join/windows) , anteriormente chamado de "canal mensal (direcionado)" ou "insider Slow".

> [!NOTE]
> Guias contextuais personalizadas funcionam somente em plataformas que dão suporte aos seguintes conjuntos de requisitos. Para saber mais sobre conjuntos de requisitos e como trabalhar com eles, confira [especificar aplicativos do Office e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).
>
> - [SharedRuntime 1,1](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a>Comportamento de guias contextuais personalizadas

A experiência do usuário para guias contextuais personalizadas segue o padrão de guias contextuais internas do Office. Estes são os princípios básicos para as guias contextuais personalizadas de posicionamento:

- Quando uma guia contextual personalizada estiver visível, ela aparecerá na extremidade direita da faixa de opções.
- Se uma ou mais guias contextuais internas e uma ou mais guias contextuais personalizadas de suplementos forem visíveis ao mesmo tempo, as guias contextuais personalizadas estarão sempre à direita de todas as guias contextuais internas.
- Se o suplemento tiver mais de uma guia contextual e houver contextos em que mais de uma esteja visível, elas aparecerão na ordem em que estão definidas no suplemento. (A direção é a mesma direção do idioma do Office; ou seja, da esquerda para a direita em idiomas da esquerda para a direita, mas da direita para a esquerda em idiomas da direita para a esquerda.) Consulte [definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como você os define.
- Se mais de um suplemento tiver uma guia contextual que seja visível em um contexto específico, elas aparecerão na ordem em que os suplementos foram iniciados.
- As guias *contextuais* personalizadas, diferente das guias principais personalizadas, não são adicionadas permanentemente à faixa de opções do aplicativo do Office. Eles estão presentes somente nos documentos do Office em que o suplemento está sendo executado.

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a>Etapas principais para incluir uma guia contextual em um suplemento

A seguir estão as principais etapas para incluir uma guia contextual personalizada em um suplemento:

1. Configure o suplemento para usar um tempo de execução compartilhado.
1. Defina a guia e os grupos e controles que aparecem nele.
1. Registre a guia contextual com o Office.
1. Especifique as circunstâncias em que a guia estará visível.

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a>Configurar o suplemento para usar um tempo de execução compartilhado

A adição de guias contextuais personalizadas exige que seu suplemento use o tempo de execução compartilhado. Para obter mais informações, consulte [configurar um suplemento para usar um tempo de execução compartilhado](../excel/configure-your-add-in-to-use-a-shared-runtime.md).

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a>Definir os grupos e controles que aparecem na guia

Ao contrário das guias principais personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com um blob JSON. O código analisa o blob em um objeto JavaScript e, em seguida, passa o objeto para o método [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) . As guias contextuais personalizadas só estão presentes em documentos nos quais o suplemento está sendo executado. Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções do aplicativo do Office quando o suplemento é instalado e permanecer presente quando outro documento é aberto. Além disso, o `requestCreateControls` método pode ser executado apenas uma vez em uma sessão do seu suplemento. Se for chamado novamente, um erro será gerado.

> [!NOTE]
> A estrutura das propriedades e subpropriedades do blob JSON (e os nomes das chaves) é quase paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no XML do manifesto.

Criaremos um exemplo de um blob JSON de guias contextual passo a passo. (O esquema completo para a guia contextual JSON está em [dynamic-ribbon.schema.js](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json). Este link pode não estar funcionando no período de visualização inicial para guias contextuais. Se o link não estiver funcionando, você poderá encontrar o rascunho mais recente do esquema em [rascunho dynamic-ribbon.schema.jsem](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Se você estiver trabalhando no Visual Studio Code, você pode usar esse arquivo para obter o IntelliSense e para validar seu JSON. Para obter mais informações, consulte [Editing JSON with Visual Studio Code-JSON schemas and Settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).


1. Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz chamadas `actions` e `tabs` . A `actions` matriz é uma especificação de todas as funções que podem ser executadas pelos controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, *até o máximo de 10*.

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, uma única ação. Adicione o seguinte como o único membro da `actions` matriz. Sobre essa marcação, observe:

    - As `id` `type` Propriedades e são obrigatórias.
    - O valor de `type` pode ser "ExecuteFunction" ou "ShowTaskpane".
    - A `functionName` propriedade é usada somente quando o valor de `type` é `ExecuteFunction` . É o nome de uma função definida no Functionfile. Para obter mais informações sobre o Functionfile, consulte [conceitos básicos para comandos de suplemento](add-in-commands.md).
    - Em uma etapa posterior, você irá mapear essa ação para um botão na guia contextual.

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. Adicione o seguinte como o único membro da `tabs` matriz. Sobre essa marcação, observe:

    - A propriedade `id` é obrigatória. Use uma ID curta e descritiva exclusiva entre todas as guias contextuais no seu suplemento.
    - A propriedade `label` é obrigatória. É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.
    - A propriedade `groups` é obrigatória. Ele define os grupos de controles que serão exibidos na guia. Deve ter pelo menos um membro e, no máximo *, 20*. (Também há limites quanto ao número de controles que você pode ter em uma guia contextual personalizada e que também restringe o número de grupos que você tem. Consulte a próxima etapa para obter mais informações.)

    > [!NOTE]
    > O objeto Tab também pode ter uma `visible` propriedade opcional que especifica se a guia estará visível imediatamente quando o suplemento for iniciado. Como as guias contextuais são normalmente ocultas até que um evento de usuário dispare sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a `visible` propriedade será definida como padrão `false` quando não estiver presente. Em uma seção posterior, mostraremos como definir a propriedade como `true` em resposta a um evento.

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. No exemplo simples em andamento, a guia contextual tem apenas um único grupo. Adicione o seguinte como o único membro da `groups` matriz. Sobre essa marcação, observe:

    - Todas as propriedades são obrigatórias.
    - A `id` propriedade deve ser exclusiva entre todos os grupos na guia. Use uma ID breve e descritiva.
    - O `label` é uma cadeia de caracteres amigável para servir como o rótulo do grupo.
    - O `icon` valor da propriedade é uma matriz de objetos que especifica os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela do aplicativo do Office.
    - O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e outros controles no grupo. Deve haver pelo menos um e *não mais de 6 em um grupo*.

    > [!IMPORTANT]
    > *O número total de controles na guia inteira não pode ser superior a 20.* Por exemplo, você poderia ter 3 grupos com 6 controles cada e um quarto grupo com 2 controles, mas não pode ter quatro grupos com 6 controles cada.  

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

1. Todos os grupos devem ter um ícone de pelo menos dois tamanhos, 32x32 PX e 80x80 px. Opcionalmente, você também pode ter ícones de tamanhos 16x16, 20x20, 24x24, 40x40, 48x48 e 64x64. O Office decide qual ícone usar com base no tamanho da faixa de opções e na janela do aplicativo do Office. Adicione os seguintes objetos à matriz de ícones. (Se a janela e os tamanhos de faixa de opções forem grandes o suficiente para que pelo menos um dos *controles* no grupo apareça, então nenhum ícone de grupo aparecerá. Por exemplo, Assista ao grupo **estilos** na faixa de opções do Word à medida que você encolhe e expande a janela do Word.) Sobre essa marcação, observe:

    - Ambas as propriedades são obrigatórias.
    - A `size` unidade de medida de propriedade é pixels. Os ícones são sempre quadrados, portanto, o número é a altura e a largura.
    - A `sourceLocation` propriedade especifica a URL completa para o ícone.

    > [!IMPORTANT]
    > Assim como você deve alterar as URLs no manifesto do suplemento quando migrar do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.

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

1. Em nosso exemplo simples em andamento, o grupo tem apenas um único botão. Adicione o seguinte objeto como o único membro da `controls` matriz. Sobre essa marcação, observe:

    - Todas as propriedades, exceto `enabled` , são obrigatórias.
    - `type` Especifica o tipo de controle. Os valores podem ser "Button", "menu" ou "MobileButton".
    - `id` pode ter até 125 caracteres. 
    - `actionId` deve ser a ID de uma ação definida na `actions` matriz. (Confira a etapa 1 desta seção.)
    - `label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.
    - `superTip` representa uma forma rica da dica de ferramenta. As `title` Propriedades e `description` são obrigatórias.
    - `icon` Especifica os ícones para o botão. Os comentários anteriores sobre o ícone de grupo aplicam-se aqui também.
    - `enabled` (opcional) especifica se o botão está habilitado quando a guia contextual aparece é iniciada. O padrão, se não estiver presente, é `true` . 

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
'{
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
      "label": "Data",
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
}'
```

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a>Registrar a guia contextual com o Office com o requestCreateControls

A guia contextual é registrada no Office chamando o método [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) . Isso geralmente é feito na função que é atribuída `Office.initialize` ou ao `Office.onReady` método. Para saber mais sobre esses métodos e inicializar o suplemento, confira [inicializar o suplemento do Office](../develop/initialize-add-in.md). No entanto, você pode chamar o método a qualquer momento após a inicialização.

> [!IMPORTANT]
> O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um suplemento. Um erro será acionado se for chamado novamente.

Apresentamos um exemplo a seguir. Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o `JSON.parse` método antes que ele possa ser passado para uma função JavaScript.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a>Especifique os contextos quando a guia estará visível com requestUpdate

Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto do suplemento. Considere um cenário em que a guia deve estar visível quando, e somente quando, um gráfico (na planilha padrão de uma pasta de trabalho do Excel) estiver ativado.

Comece atribuindo manipuladores. Isso geralmente é feito no `Office.onReady` método como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos `onActivated` `onDeactivated` eventos e de todos os gráficos da planilha.

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string.
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

Em seguida, defina os manipuladores. Veja a seguir um exemplo simples de um `showDataTab` , mas consulte [tratamento de erros](#error-handling) posteriormente neste artigo para obter uma versão mais robusta da função. Sobre este código, observe:

- O Office controla quando atualiza o estado da faixa de opções. O método  [Office. Ribbon. requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) enfileira uma solicitação para atualizar. O método resolverá o `Promise` objeto assim que ele enfileirar a solicitação, não quando a faixa de opções realmente for atualizada.
- O parâmetro para o `requestUpdate` método é um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a Tabulação por sua ID *exatamente conforme especificado no JSON* e (2) especifica a visibilidade da guia.
- Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar outros objetos Tab à `tabs` matriz.

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

O manipulador para ocultar a guia é quase idêntico, exceto pelo fato de que ela define a `visible` propriedade de volta para `false` .

A biblioteca JavaScript do Office também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto. A seguir está a `showDataTab` função no TypeScript e utiliza esses tipos.

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Alternar a visibilidade da guia e o status habilitado de um botão ao mesmo tempo

O `requestUpdate` método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada ou em uma guia principal personalizada. Para obter detalhes sobre isso, consulte [habilitar e desabilitar comandos de suplemento](disable-add-in-commands.md). Pode haver cenários nos quais você deseja alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo. Você pode fazer isso com uma única chamada de `requestUpdate` . A seguir está um exemplo no qual um botão em uma guia principal está habilitado ao mesmo tempo em que uma guia contextual é torna visível.

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
                controls: [
                {
                    id: "MyButton",
                    enabled: true
                }
            ]}
        ]});
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
                controls: [
                    {
                        id: "MyButton",
                        enabled: true
                    }
                ]
            }
        ]});
}
```

## <a name="error-handling"></a>Tratamento de erros

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
