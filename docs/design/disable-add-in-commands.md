---
title: Habilitar e Desabilitar Comandos de Suplemento
description: Aprenda a alterar o status habilitado ou desabilitado dos botões da faixa de opções personalizados e itens de menu no seu Suplemento da Web do Office.
ms.date: 04/11/2020
localization_priority: Priority
ms.openlocfilehash: b223949c5665f0e523a6003ccb4e3f29cc3f017c
ms.sourcegitcommit: 79c55e59294e220bd21a5006080f72acf3ec0a3f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2020
ms.locfileid: "43581922"
---
# <a name="enable-and-disable-add-in-commands-preview"></a><span data-ttu-id="91fc8-103">Habilitar e Desabilitar Comandos de Suplemento (visualização)</span><span class="sxs-lookup"><span data-stu-id="91fc8-103">Enable and Disable Add-in Commands (preview)</span></span>

<span data-ttu-id="91fc8-104">Quando alguma funcionalidade do seu suplemento deve estar disponível apenas em determinados contextos, você pode habilitar ou desabilitar programaticamente seus Comandos de Suplemento personalizados.</span><span class="sxs-lookup"><span data-stu-id="91fc8-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="91fc8-105">Por exemplo, uma função que altera o cabeçalho de uma tabela só deve ser ativada quando o cursor estiver em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="91fc8-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="91fc8-106">Você também pode especificar se o comando será ativado ou desativado quando o aplicativo host do Office for aberto.</span><span class="sxs-lookup"><span data-stu-id="91fc8-106">You can also specify whether the command is enabled or disabled when the Office host application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="91fc8-107">Este artigo pressupõe que você esteja familiarizado com a seguinte documentação.</span><span class="sxs-lookup"><span data-stu-id="91fc8-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="91fc8-108">Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).</span><span class="sxs-lookup"><span data-stu-id="91fc8-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> [<span data-ttu-id="91fc8-109">Conceitos básicos dos Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="91fc8-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="preview-status"></a><span data-ttu-id="91fc8-110">Status da visualização</span><span class="sxs-lookup"><span data-stu-id="91fc8-110">Preview status</span></span>

<span data-ttu-id="91fc8-111">As APIs descritas neste artigo estão em visualização e atualmente estão disponíveis apenas no Excel.</span><span class="sxs-lookup"><span data-stu-id="91fc8-111">The APIs described in this article are in preview and are currently only available in Excel.</span></span>

> [!NOTE]
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="rules-and-gotchas"></a><span data-ttu-id="91fc8-112">Regras e dicas</span><span class="sxs-lookup"><span data-stu-id="91fc8-112">Rules and gotchas</span></span>

### <a name="single-line-ribbon-in-office-on-the-web"></a><span data-ttu-id="91fc8-113">Faixa de opções de linha única no Office na Web</span><span class="sxs-lookup"><span data-stu-id="91fc8-113">Single-line ribbon in Office on the web</span></span>

<span data-ttu-id="91fc8-114">No Office na Web, as APIs e a marcação de manifesto descritas neste artigo afetam apenas a faixa de opções de linha única.</span><span class="sxs-lookup"><span data-stu-id="91fc8-114">In Office on the web, the APIs and manifest markup described in this article only affect the single-line ribbon.</span></span> <span data-ttu-id="91fc8-115">Elas não têm efeito na faixa de opções de várias linhas.</span><span class="sxs-lookup"><span data-stu-id="91fc8-115">They have no effect on the multiline ribbon.</span></span> <span data-ttu-id="91fc8-116">Eles afetam as duas faixas de opções do Office para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="91fc8-116">They affect both ribbons for desktop Office.</span></span> <span data-ttu-id="91fc8-117">Para obter mais informações sobre as duas faixas de opções, confira [Usar a faixa de opções simplificada](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span><span class="sxs-lookup"><span data-stu-id="91fc8-117">For more information about the two ribbons, see [Use the simplified ribbon](https://support.office.com/article/Use-the-Simplified-Ribbon-44bef9c3-295d-4092-b7f0-f471fa629a98).</span></span>

### <a name="shared-runtime-required"></a><span data-ttu-id="91fc8-118">Tempo de execução compartilhado necessário</span><span class="sxs-lookup"><span data-stu-id="91fc8-118">Shared runtime required</span></span>

<span data-ttu-id="91fc8-119">As APIs e a marcação de manifesto descritas neste artigo exigem que o manifesto do suplemento especifique que ele deve usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="91fc8-119">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="91fc8-120">Para fazer isso, execute as seguintes etapas.</span><span class="sxs-lookup"><span data-stu-id="91fc8-120">To do this take the following steps.</span></span>

1. <span data-ttu-id="91fc8-121">No elemento [Runtimes](../reference/manifest/runtimes.md) no manifesto, adicione o seguinte elemento filho: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-121">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="91fc8-122">(Se ainda não houver um elemento `<Runtimes>` no manifesto, crie-o como o primeiro filho abaixo do elemento `<Host>` na seção `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="91fc8-122">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="91fc8-123">Na seção [Resources.Urls](../reference/manifest/resources.md) do manifesto, adicione o seguinte elemento filho: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, onde `{MyDomain}` é o domínio do suplemento e `{path-to-start-page}` o caminho da página inicial do suplemento; por exemplo: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-123">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="91fc8-124">Dependendo do seu suplemento conter um painel de tarefas, um arquivo de função ou uma função personalizada do Excel, você deve executar uma ou mais das três etapas a seguir:</span><span class="sxs-lookup"><span data-stu-id="91fc8-124">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="91fc8-125">Se o suplemento contiver um painel de tarefas, defina o `resid` atributo do elemento [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do elemento `<Runtime>` na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-125">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="91fc8-126">O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-126">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="91fc8-127">Se o suplemento contiver uma função personalizada do Excel, defina o `resid` atributo do elemento [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-127">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="91fc8-128">O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-128">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="91fc8-129">Se o suplemento contiver um arquivo de função, defina o `resid` atributo do elemento [FunctionFile](../reference/manifest/functionfile.md) para exatamente a mesma série que você usou para o `resid`do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-129">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="91fc8-130">O elemento deve ficar assim: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-130">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="91fc8-131">Defina o estado padrão como desabilitado</span><span class="sxs-lookup"><span data-stu-id="91fc8-131">Set the default state to disabled</span></span>

<span data-ttu-id="91fc8-132">Por padrão, qualquer comando de suplemento é habilitado quando o aplicativo do Office é iniciado.</span><span class="sxs-lookup"><span data-stu-id="91fc8-132">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="91fc8-133">Se você deseja que um botão ou item de menu personalizado esteja desabilitado quando o aplicativo do Office for iniciado, especifique isso no manifesto.</span><span class="sxs-lookup"><span data-stu-id="91fc8-133">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="91fc8-134">Basta adicionar um elemento [Enabled](../reference/manifest/enabled.md) (com o valor `false`) imediatamente *abaixo* (não dentro) do elemento [Ação](../reference/manifest/action.md) na declaração do controle.</span><span class="sxs-lookup"><span data-stu-id="91fc8-134">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="91fc8-135">Veja a estrutura básica a seguir:</span><span class="sxs-lookup"><span data-stu-id="91fc8-135">The following shows the basic structure:</span></span>

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

## <a name="change-the-state-programmatically"></a><span data-ttu-id="91fc8-136">Alterar o estado programaticamente</span><span class="sxs-lookup"><span data-stu-id="91fc8-136">Change the state programmatically</span></span>

<span data-ttu-id="91fc8-137">As etapas essenciais para alterar o status habilitado de um Comando de Suplemento são:</span><span class="sxs-lookup"><span data-stu-id="91fc8-137">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="91fc8-138">Criar um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifique o comando e sua guia pai por seus IDs, conforme especificado no manifesto; e (2) especifica o estado habilitado ou desabilitado do comando.</span><span class="sxs-lookup"><span data-stu-id="91fc8-138">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent tab, by their IDs as specified in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="91fc8-139">Passe o objeto **RibbonUpdaterData** para o método [OfficeRuntime.Ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="91fc8-139">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js#requestupdate-input-) method.</span></span>

<span data-ttu-id="91fc8-140">Apresentamos um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="91fc8-140">The following is a simple example.</span></span> <span data-ttu-id="91fc8-141">Observe que "MyButton" e "OfficeAddinTab1" são copiados do manifesto.</span><span class="sxs-lookup"><span data-stu-id="91fc8-141">Note that "MyButton" and "OfficeAddinTab1" are copied from the manifest.</span></span>

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

<span data-ttu-id="91fc8-142">Também fornecemos várias interfaces (tipos) para facilitar a construção do objeto **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="91fc8-142">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="91fc8-143">Veja a seguir o exemplo equivalente no TypeScript, que faz uso desses tipos.</span><span class="sxs-lookup"><span data-stu-id="91fc8-143">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentTab: Tab = {id: "OfficeAddinTab1", controls: [button]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="91fc8-144">O Office controla quando atualiza o estado da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="91fc8-144">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="91fc8-145">O método **requestUpdate()** adiciona uma solicitação para atualização à fila de espera.</span><span class="sxs-lookup"><span data-stu-id="91fc8-145">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="91fc8-146">O método resolverá o objeto Promise assim que a solicitação estiver na fila, não quando a faixa de opções for de fato atualizada.</span><span class="sxs-lookup"><span data-stu-id="91fc8-146">The method will resolve the Promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="91fc8-147">Alterar o estado em resposta a um evento</span><span class="sxs-lookup"><span data-stu-id="91fc8-147">Change the state in response to an event</span></span>

<span data-ttu-id="91fc8-148">Um cenário comum em que o estado da faixa de opções deve mudar é quando um evento iniciado pelo usuário altera o contexto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="91fc8-148">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="91fc8-149">Considere um cenário em que um botão deve ser ativado quando e somente quando um gráfico é ativado.</span><span class="sxs-lookup"><span data-stu-id="91fc8-149">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="91fc8-150">A primeira etapa é definir o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto como `false`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-150">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="91fc8-151">Veja um exemplo acima.</span><span class="sxs-lookup"><span data-stu-id="91fc8-151">See above for an example.</span></span>

<span data-ttu-id="91fc8-152">Segundo, atribua manipuladores.</span><span class="sxs-lookup"><span data-stu-id="91fc8-152">Second, assign handlers.</span></span> <span data-ttu-id="91fc8-153">Isso geralmente é feito no método **Office.onReady**, como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos eventos **onActivated** e **onDeactivated** de todos os gráficos da planilha.</span><span class="sxs-lookup"><span data-stu-id="91fc8-153">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="91fc8-154">Terceiro, defina o manipulador `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-154">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="91fc8-155">A seguir, é apresentado um exemplo simples, mas consulte **Prática recomendada: Teste se há erros de status do controle** abaixo para obter uma maneira mais robusta de alterar o status de um controle.</span><span class="sxs-lookup"><span data-stu-id="91fc8-155">The following is a simple example, but see **Best practice: Test for control status errors** below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: true};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="91fc8-156">Quarto, defina o manipulador `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-156">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="91fc8-157">Seria idêntico a `enableChartFormat`, exceto que a propriedade **enabled** do objeto button seria configurada como `false`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-157">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="91fc8-158">Prática recomendada: Teste se há erros de status do controle</span><span class="sxs-lookup"><span data-stu-id="91fc8-158">Best practice: Test for control status errors</span></span>

<span data-ttu-id="91fc8-159">Em algumas circunstâncias, a faixa de opções não é redesenhada após `requestUpdate` ser chamado, portanto, o status clicável do controle não muda.</span><span class="sxs-lookup"><span data-stu-id="91fc8-159">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="91fc8-160">Por esse motivo, é uma prática recomendada para o suplemento acompanhar o status de seus controles.</span><span class="sxs-lookup"><span data-stu-id="91fc8-160">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="91fc8-161">O suplemento deve estar em conformidade com estas regras:</span><span class="sxs-lookup"><span data-stu-id="91fc8-161">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="91fc8-162">Sempre que `requestUpdate` é chamado, o código deve registrar o estado pretendido dos botões e itens de menu personalizados.</span><span class="sxs-lookup"><span data-stu-id="91fc8-162">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="91fc8-163">Quando um controle personalizado é clicado, o primeiro código no manipulador deve verificar se o botão deveria ter sido clicável.</span><span class="sxs-lookup"><span data-stu-id="91fc8-163">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="91fc8-164">Se não deveria ter sido, o código deve relatar ou registrar um erro e tentar novamente definir os botões no estado pretendido.</span><span class="sxs-lookup"><span data-stu-id="91fc8-164">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="91fc8-165">O exemplo a seguir mostra uma função que desativa um botão e registra o status do botão.</span><span class="sxs-lookup"><span data-stu-id="91fc8-165">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="91fc8-166">Observe que `chartFormatButtonEnabled` é uma variável booleana global inicializada com o mesmo valor que o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto.</span><span class="sxs-lookup"><span data-stu-id="91fc8-166">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {id: "ChartFormatButton", enabled: false};
    var parentTab = {id: "CustomChartTab", controls: [button]};
    var ribbonUpdater = {tabs: [parentTab]};
    await Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="91fc8-167">O exemplo a seguir mostra como o manipulador do botão testa um estado incorreto do botão.</span><span class="sxs-lookup"><span data-stu-id="91fc8-167">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="91fc8-168">Observe que `reportError` é uma função que mostra ou registra um erro.</span><span class="sxs-lookup"><span data-stu-id="91fc8-168">Note that `reportError` is a function that shows or logs an error.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="91fc8-169">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="91fc8-169">Error handling</span></span>

<span data-ttu-id="91fc8-170">Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="91fc8-170">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="91fc8-171">Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="91fc8-171">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="91fc8-172">Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="91fc8-172">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="91fc8-173">Veja um exemplo de como lidar com esse erro a seguir.</span><span class="sxs-lookup"><span data-stu-id="91fc8-173">The following is an example of how to handle this error.</span></span> <span data-ttu-id="91fc8-174">Nesse caso, o método `reportError` exibe o erro para o usuário.</span><span class="sxs-lookup"><span data-stu-id="91fc8-174">In this case, the `reportError` method displays the error to the user.</span></span>

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

## <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="91fc8-175">Teste se há suporte à plataforma com conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="91fc8-175">Test for platform support with requirement sets</span></span>

<span data-ttu-id="91fc8-p122">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="91fc8-p122">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="91fc8-179">As APIs de ativação/desativação requerem suporte do seguinte conjunto de requisitos:</span><span class="sxs-lookup"><span data-stu-id="91fc8-179">The enable/disable APIs require support of the following requirement set:</span></span>

- [<span data-ttu-id="91fc8-180">AddinCommands 1.1</span><span class="sxs-lookup"><span data-stu-id="91fc8-180">AddinCommands 1.1</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
