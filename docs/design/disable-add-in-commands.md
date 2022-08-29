---
title: Habilitar e Desabilitar Comandos de Suplemento
description: Aprenda a alterar o status habilitado ou desabilitado dos botões da faixa de opções personalizados e itens de menu no seu Suplemento da Web do Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 502c9247a6c63775c562dab7479e0ca926f14154
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423045"
---
# <a name="enable-and-disable-add-in-commands"></a>Habilitar e Desabilitar Comandos de Suplemento

Quando alguma funcionalidade do seu suplemento deve estar disponível apenas em determinados contextos, você pode habilitar ou desabilitar programaticamente seus Comandos de Suplemento personalizados. Por exemplo, uma função que altera o cabeçalho de uma tabela só deve ser ativada quando o cursor estiver em uma tabela.

Você também pode especificar se o comando está habilitado ou desabilitado quando o aplicativo cliente do Office é aberto.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Somente suporte a aplicativos e plataformas do Office

As APIs descritas neste artigo só estão disponíveis no Excel, no PowerPoint e no Word.

### <a name="test-for-platform-support-with-requirement-sets"></a>Teste se há suporte à plataforma com conjuntos de requisitos

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de runtime para determinar se uma combinação de aplicativo e plataforma do Office dá suporte a APIs de que um suplemento precisa. Para obter mais informações, consulte [versões do Office e conjuntos de requisitos](../develop/office-versions-and-requirement-sets.md).

As APIs de habilitação/desabilitação pertencem ao conjunto de requisitos [RibbonApi 1.1](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets) .

> [!NOTE]
> O **conjunto de requisitos RibbonApi 1.1** ainda não tem suporte no manifesto, portanto, não é possível especificá-lo na seção do **\<Requirements\>** manifesto. Para testar o suporte, seu código deve chamar `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`. Se, *e somente se*, essa chamada retornar `true`, seu código poderá chamar as APIs de habilitação/desabilitação. Se a chamada de `isSetSupported` retorno for `false`retornada, todos os comandos de suplemento personalizados serão habilitados o tempo todo. Você deve projetar seu suplemento de produção e todas as instruções no aplicativo para levar em conta como ele funcionará quando o conjunto de requisitos **RibbonApi 1.1** não tiver suporte. Para obter mais informações e exemplos de uso, consulte Especificar aplicativos `isSetSupported`do Office e requisitos de [API](../develop/specify-office-hosts-and-api-requirements.md), especialmente verificações de runtime para suporte ao [método e ao conjunto de requisitos](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). (A seção [Especificar quais versões e plataformas do Office](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in) podem hospedar seu suplemento desse artigo não se aplica à Faixa de Opções 1.1.)

## <a name="shared-runtime-required"></a>Tempo de execução compartilhado necessário

As APIs e a marcação de manifesto descritas neste artigo exigem que o manifesto do suplemento especifique que ele deve usar um [runtime compartilhado](../testing/runtimes.md#shared-runtime). Para fazer isso, execute as etapas a seguir.

1. No elemento [Runtimes](/javascript/api/manifest/runtimes) no manifesto, adicione o seguinte elemento filho: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Se ainda não houver um **\<Runtimes\>** elemento no manifesto, crie-o como o primeiro filho **\<Host\>** sob o elemento na **\<VersionOverrides\>** seção.)
2. Na seção [Resources.Urls](/javascript/api/manifest/resources) do manifesto, adicione o seguinte elemento filho: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, onde `{MyDomain}` é o domínio do suplemento e `{path-to-start-page}` o caminho da página inicial do suplemento; por exemplo: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. Dependendo se o suplemento contém um painel de tarefas, um arquivo de função ou uma função personalizada do Excel, você deve executar uma ou mais das três etapas a seguir.

    - Se o suplemento contiver um painel de tarefas, defina o `resid` atributo da [Ação](/javascript/api/manifest/action).[ Elemento SourceLocation](/javascript/api/manifest/sourcelocation) para exatamente a mesma cadeia de caracteres `resid` **\<Runtime\>** usada para o elemento na etapa 1; por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Se o suplemento contiver uma função personalizada do Excel, defina o `resid` atributo da [Página](/javascript/api/manifest/page).[ Elemento SourceLocation](/javascript/api/manifest/sourcelocation) exatamente a mesma cadeia de caracteres `resid` **\<Runtime\>** usada para o elemento na etapa 1; por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Se o suplemento contiver um arquivo de função, `resid` defina o atributo do elemento [FunctionFile](/javascript/api/manifest/functionfile) `resid` **\<Runtime\>** como exatamente a mesma cadeia de caracteres usada para o elemento na etapa 1; por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Defina o estado padrão como desabilitado

Por padrão, qualquer comando de suplemento é habilitado quando o aplicativo do Office é iniciado. Se você deseja que um botão ou item de menu personalizado esteja desabilitado quando o aplicativo do Office for iniciado, especifique isso no manifesto. Basta adicionar um elemento [Enabled](/javascript/api/manifest/enabled) (com o valor `false`) imediatamente *abaixo* (não dentro) do elemento [Ação](/javascript/api/manifest/action) na declaração do controle. O exemplo a seguir mostra a estrutura básica.

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
                <Control ... id="Contoso.MyButton3">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Alterar o estado programaticamente

As etapas essenciais para alterar o status habilitado de um Comando de Suplemento são:

1. Crie um [objeto RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica o comando e seu grupo pai e guia, por suas IDs, conforme declarado no manifesto; e (2) especifica o estado habilitado ou desabilitado do comando.
2. Passe o objeto **RibbonUpdaterData** para o método [OfficeRuntime.Ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#office-office-ribbon-requestupdate-member(1)).

Apresentamos um exemplo simples a seguir. Observe que "MyButton", "OfficeAddinTab1" e "CustomGroup111" são copiados do manifesto.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
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
            }
        ]
    });
}
```

Também fornecemos várias interfaces (tipos) para facilitar a construção do objeto **RibbonUpdateData**. Veja a seguir o exemplo equivalente no TypeScript, que faz uso desses tipos.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Você pode `await` chamar **requestUpdate()** se a função pai for assíncrona, mas observe que o aplicativo do Office controla quando ele atualiza o estado da faixa de opções. O método **requestUpdate()** adiciona uma solicitação para atualização à fila de espera. O método resolverá o objeto promise assim que tiver enfileirado a solicitação, não quando a faixa de opções realmente for atualizada.

## <a name="change-the-state-in-response-to-an-event"></a>Alterar o estado em resposta a um evento

Um cenário comum em que o estado da faixa de opções deve mudar é quando um evento iniciado pelo usuário altera o contexto do suplemento.

Considere um cenário em que um botão deve ser ativado quando e somente quando um gráfico é ativado. A primeira etapa é definir o elemento [Enabled](/javascript/api/manifest/enabled) para o botão no manifesto como `false`. Veja um exemplo acima.

Segundo, atribua manipuladores. Isso normalmente é feito na função **Office.onReady** , como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos eventos **onActivated** e **onDeactivated** de todos os gráficos na planilha.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        const charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

Terceiro, defina o manipulador `enableChartFormat`. A seguir, é apresentado um exemplo simples, mas consulte [Prática recomendada: Teste se há erros de status do controle](#best-practice-test-for-control-status-errors) abaixo para obter uma maneira mais robusta de alterar o status de um controle.

```javascript
function enableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Quarto, defina o manipulador `disableChartFormat`. Seria idêntico a `enableChartFormat`, exceto que a propriedade **enabled** do objeto button seria configurada como `false`.

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a>Alternar a visibilidade da guia e o status habilitado de um botão ao mesmo tempo

O **método requestUpdate** também é usado para alternar a visibilidade de uma guia contextual personalizada. Para obter detalhes sobre esse e o código de exemplo, consulte [Criar guias contextuais personalizadas em Suplementos do Office](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).

## <a name="best-practice-test-for-control-status-errors"></a>Prática recomendada: Teste se há erros de status do controle

Em algumas circunstâncias, a faixa de opções não é redesenhada após `requestUpdate` ser chamado, portanto, o status clicável do controle não muda. Por esse motivo, é uma prática recomendada para o suplemento acompanhar o status de seus controles. O suplemento deve estar em conformidade com as regras a seguir.

1. Sempre que `requestUpdate` é chamado, o código deve registrar o estado pretendido dos botões e itens de menu personalizados.
2. Quando um controle personalizado é clicado, o primeiro código no manipulador deve verificar se o botão deveria ter sido clicável. Se não deveria ter sido, o código deve relatar ou registrar um erro e tentar novamente definir os botões no estado pretendido.

O exemplo a seguir mostra uma função que desativa um botão e registra o status do botão. Observe que `chartFormatButtonEnabled` é uma variável booleana global inicializada com o mesmo valor que o elemento [Enabled](/javascript/api/manifest/enabled) para o botão no manifesto.

```javascript
function disableChartFormat() {
    const button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    const parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    const parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    const ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

O exemplo a seguir mostra como o manipulador do botão testa um estado incorreto do botão. Observe que `reportError` é uma função que mostra ou registra um erro.

```javascript
function chartFormatButtonHandler() {
    if (chartFormatButtonEnabled) {

        // Do work here

    } else {
        // Report the error and try again to disable.
        reportError("That action is not possible at this time.");
        disableChartFormat();
    }
}
```

## <a name="error-handling"></a>Tratamento de erros

Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro. Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto. Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`. Veja um exemplo de como lidar com esse erro a seguir. Nesse caso, o método `reportError` exibe o erro para o usuário.

```javascript
function disableChartFormat() {
    try {
        const button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        const parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        const parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        const ribbonUpdater = {tabs: [parentTab]};
        Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```
