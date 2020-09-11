---
title: Habilitar e Desabilitar Comandos de Suplemento
description: Aprenda a alterar o status habilitado ou desabilitado dos botões da faixa de opções personalizados e itens de menu no seu Suplemento da Web do Office.
ms.date: 08/26/2020
localization_priority: Normal
ms.openlocfilehash: fac62b20dc67db591ba2de73f96526b8a3dfdf9e
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430412"
---
# <a name="enable-and-disable-add-in-commands"></a>Habilitar e Desabilitar Comandos de Suplemento

Quando alguma funcionalidade do seu suplemento deve estar disponível apenas em determinados contextos, você pode habilitar ou desabilitar programaticamente seus Comandos de Suplemento personalizados. Por exemplo, uma função que altera o cabeçalho de uma tabela só deve ser ativada quando o cursor estiver em uma tabela.

Você também pode especificar se o comando está habilitado ou desabilitado quando o aplicativo cliente do Office é aberto.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com a seguinte documentação. Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).
>
> - [Conceitos básicos dos Comandos de Suplemento](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a>Suporte apenas a aplicativos e plataformas do Office

As APIs descritas neste artigo estão disponíveis apenas no Excel e apenas no Office no Windows e no Mac.

### <a name="test-for-platform-support-with-requirement-sets"></a>Teste se há suporte à plataforma com conjuntos de requisitos

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se uma combinação de aplicativos e plataformas do Office oferece suporte a APIs necessárias para um suplemento. Para obter mais informações, consulte [versões do Office e conjuntos de requisitos](../develop/office-versions-and-requirement-sets.md).

As APIs Enable/Disable pertencem ao conjunto de requisitos [RibbonApi 1,1](../reference/requirement-sets/ribbon-api-requirement-sets.md) .

> [!NOTE]
> O conjunto de requisitos **RibbonApi 1,1** ainda não tem suporte no manifesto, portanto, você não pode especificá-lo na seção do manifesto `<Requirements>` . Para testar o suporte, seu código deve chamar `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` . Se, *e somente se*, essa chamada retornar `true` , seu código poderá chamar as APIs habilitar/desabilitar. Se a chamada de `isSetSupported` Devoluções `false` , todos os comandos de suplemento personalizados são habilitados todo o tempo. Você deve projetar seu suplemento de produção e quaisquer instruções no aplicativo para considerar como funcionará quando o conjunto de requisitos **RibbonApi 1,1** não for suportado. Para obter mais informações e exemplos de como usar o `isSetSupported` , consulte [especificar aplicativos do Office e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md), principalmente [usar verificações de tempo de execução em seu código JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). (A seção [define o elemento requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) desse artigo não se aplica à faixa de opções 1,1.)

## <a name="shared-runtime-required"></a>Tempo de execução compartilhado necessário

As APIs e a marcação de manifesto descritas neste artigo exigem que o manifesto do suplemento especifique que ele deve usar um tempo de execução compartilhado. Para fazer isso, execute as seguintes etapas.

1. No elemento [Runtimes](../reference/manifest/runtimes.md) no manifesto, adicione o seguinte elemento filho: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`. (Se ainda não houver um elemento `<Runtimes>` no manifesto, crie-o como o primeiro filho abaixo do elemento `<Host>` na seção `VersionOverrides`.)
2. Na seção [Resources.Urls](../reference/manifest/resources.md) do manifesto, adicione o seguinte elemento filho: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, onde `{MyDomain}` é o domínio do suplemento e `{path-to-start-page}` o caminho da página inicial do suplemento; por exemplo: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.
3. Dependendo do seu suplemento conter um painel de tarefas, um arquivo de função ou uma função personalizada do Excel, você deve executar uma ou mais das três etapas a seguir:

    - Se o suplemento contiver um painel de tarefas, defina o `resid` atributo do elemento [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do elemento `<Runtime>` na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Se o suplemento contiver uma função personalizada do Excel, defina o `resid` atributo do elemento [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.
    - Se o suplemento contiver um arquivo de função, defina o `resid` atributo do elemento [FunctionFile](../reference/manifest/functionfile.md) para exatamente a mesma série que você usou para o `resid`do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`. O elemento deve ficar assim: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.

## <a name="set-the-default-state-to-disabled"></a>Defina o estado padrão como desabilitado

Por padrão, qualquer comando de suplemento é habilitado quando o aplicativo do Office é iniciado. Se você deseja que um botão ou item de menu personalizado esteja desabilitado quando o aplicativo do Office for iniciado, especifique isso no manifesto. Basta adicionar um elemento [Enabled](../reference/manifest/enabled.md) (com o valor `false`) imediatamente *abaixo* (não dentro) do elemento [Ação](../reference/manifest/action.md) na declaração do controle. Veja a estrutura básica a seguir:

```xml
<OfficeApp ...>
  ...
  <VersionOverrides ...>
    ...
    <Hosts>
      <Host ...>
        ...
        <DesktopFormFactor>
          <ExtensionPoint ...>
            <CustomTab ...>
              ...
              <Group ...>
                ...
                <Control ... id="MyButton">
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a>Alterar o estado programaticamente

As etapas essenciais para alterar o status habilitado de um Comando de Suplemento são:

1. Criar um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifique o comando e sua guia pai por seus IDs, conforme especificado no manifesto; e (2) especifica o estado habilitado ou desabilitado do comando.
2. Passe o objeto **RibbonUpdaterData** para o método [OfficeRuntime.Ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-).

Apresentamos um exemplo simples a seguir. Observe que "MyButton" e "OfficeAddinTab1" são copiados do manifesto.

```javascript
function enableButton() {
    Office.ribbon.requestUpdate({
        tabs: [
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

Também fornecemos várias interfaces (tipos) para facilitar a construção do objeto **RibbonUpdateData**. Veja a seguir o exemplo equivalente no TypeScript, que faz uso desses tipos.

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

O Office controla quando atualiza o estado da faixa de opções. O método **requestUpdate()** adiciona uma solicitação para atualização à fila de espera. O método resolverá o objeto Promise assim que a solicitação estiver na fila, não quando a faixa de opções for de fato atualizada.

## <a name="change-the-state-in-response-to-an-event"></a>Alterar o estado em resposta a um evento

Um cenário comum em que o estado da faixa de opções deve mudar é quando um evento iniciado pelo usuário altera o contexto do suplemento.

Considere um cenário em que um botão deve ser ativado quando e somente quando um gráfico é ativado. A primeira etapa é definir o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto como `false`. Veja um exemplo acima.

Segundo, atribua manipuladores. Isso geralmente é feito no método **Office.onReady**, como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos eventos **onActivated** e **onDeactivated** de todos os gráficos da planilha.

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

Terceiro, defina o manipulador `enableChartFormat`. A seguir, é apresentado um exemplo simples, mas consulte [Prática recomendada: Teste se há erros de status do controle](#best-practice-test-for-control-status-errors) abaixo para obter uma maneira mais robusta de alterar o status de um controle.

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

Quarto, defina o manipulador `disableChartFormat`. Seria idêntico a `enableChartFormat`, exceto que a propriedade **enabled** do objeto button seria configurada como `false`.

## <a name="best-practice-test-for-control-status-errors"></a>Prática recomendada: Teste se há erros de status do controle

Em algumas circunstâncias, a faixa de opções não é redesenhada após `requestUpdate` ser chamado, portanto, o status clicável do controle não muda. Por esse motivo, é uma prática recomendada para o suplemento acompanhar o status de seus controles. O suplemento deve estar em conformidade com estas regras:

1. Sempre que `requestUpdate` é chamado, o código deve registrar o estado pretendido dos botões e itens de menu personalizados.
2. Quando um controle personalizado é clicado, o primeiro código no manipulador deve verificar se o botão deveria ter sido clicável. Se não deveria ter sido, o código deve relatar ou registrar um erro e tentar novamente definir os botões no estado pretendido.

O exemplo a seguir mostra uma função que desativa um botão e registra o status do botão. Observe que `chartFormatButtonEnabled` é uma variável booleana global inicializada com o mesmo valor que o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto.

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

O exemplo a seguir mostra como o manipulador do botão testa um estado incorreto do botão. Observe que `reportError` é uma função que mostra ou registra um erro.

```javascript
function chartFormatButtonHandler() {
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
        var button = {id: "ChartFormatButton", enabled: false};
        var parentTab = {id: "CustomChartTab", controls: [button]};
        var ribbonUpdater = {tabs: [parentTab]};
        await Office.ribbon.requestUpdate(ribbonUpdater);

        chartFormatButtonEnabled = false;
    }
    catch(error) {
        if (error.code == "HostRestartNeeded"){
            reportError("Contoso Awesome Add-in has been upgraded. Please save your work, close the Office application, and restart it.");
        }
    }
}
```

## <a name="test-for-platform-support-with-requirement-sets"></a>Teste se há suporte à plataforma com conjuntos de requisitos

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../develop/office-versions-and-requirement-sets.md).

As APIs de ativação/desativação requerem suporte do seguinte conjunto de requisitos:

- [RibbonApi 1,1](../reference/requirement-sets/ribbon-api-requirement-sets.md)

