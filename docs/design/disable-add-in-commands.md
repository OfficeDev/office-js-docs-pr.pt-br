---
title: Habilitar e Desabilitar Comandos de Suplemento
description: Aprenda a alterar o status habilitado ou desabilitado dos botões da faixa de opções personalizados e itens de menu no seu Suplemento da Web do Office.
ms.date: 04/30/2021
localization_priority: Normal
ms.openlocfilehash: 9690850b2206c09b99dfc826dae1ecef915d5a04
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330154"
---
# <a name="enable-and-disable-add-in-commands"></a><span data-ttu-id="165f7-103">Habilitar e Desabilitar Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="165f7-103">Enable and Disable Add-in Commands</span></span>

<span data-ttu-id="165f7-104">Quando alguma funcionalidade do seu suplemento deve estar disponível apenas em determinados contextos, você pode habilitar ou desabilitar programaticamente seus Comandos de Suplemento personalizados.</span><span class="sxs-lookup"><span data-stu-id="165f7-104">When some functionality in your add-in should only be available in certain contexts, you can programmatically enable or disable your custom Add-in Commands.</span></span> <span data-ttu-id="165f7-105">Por exemplo, uma função que altera o cabeçalho de uma tabela só deve ser ativada quando o cursor estiver em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="165f7-105">For example, a function that changes the header of a table should only be enabled when the cursor is in a table.</span></span>

<span data-ttu-id="165f7-106">Você também pode especificar se o comando está habilitado ou desabilitado quando o aplicativo Office cliente é aberto.</span><span class="sxs-lookup"><span data-stu-id="165f7-106">You can also specify whether the command is enabled or disabled when the Office client application opens.</span></span>

> [!NOTE]
> <span data-ttu-id="165f7-107">Este artigo pressupõe que você esteja familiarizado com a seguinte documentação.</span><span class="sxs-lookup"><span data-stu-id="165f7-107">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="165f7-108">Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).</span><span class="sxs-lookup"><span data-stu-id="165f7-108">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="165f7-109">Conceitos básicos dos Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="165f7-109">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

## <a name="office-application-and-platform-support-only"></a><span data-ttu-id="165f7-110">Office suporte somente a aplicativos e plataformas</span><span class="sxs-lookup"><span data-stu-id="165f7-110">Office application and platform support only</span></span>

<span data-ttu-id="165f7-111">As APIs descritas neste artigo estão disponíveis apenas Excel em todas as plataformas e PowerPoint na Web.</span><span class="sxs-lookup"><span data-stu-id="165f7-111">The APIs described in this article are only available in Excel on all platforms and in PowerPoint on the web.</span></span>

### <a name="test-for-platform-support-with-requirement-sets"></a><span data-ttu-id="165f7-112">Teste se há suporte à plataforma com conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="165f7-112">Test for platform support with requirement sets</span></span>

<span data-ttu-id="165f7-113">Os conjuntos de requisitos são grupos nomeados de membros da API.</span><span class="sxs-lookup"><span data-stu-id="165f7-113">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="165f7-114">Office Os complementos usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se uma combinação Office aplicativo e plataforma oferece suporte a APIs que um complemento precisa.</span><span class="sxs-lookup"><span data-stu-id="165f7-114">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application and platform combination supports APIs that an add-in needs.</span></span> <span data-ttu-id="165f7-115">Para obter mais informações, [consulte Office versões e conjuntos de requisitos.](../develop/office-versions-and-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="165f7-115">For more information, see [Office versions and requirement sets](../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="165f7-116">As APIs enable/disable pertencem ao conjunto de [requisitos RibbonApi 1.1.](../reference/requirement-sets/ribbon-api-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="165f7-116">The enable/disable APIs belong to the [RibbonApi 1.1](../reference/requirement-sets/ribbon-api-requirement-sets.md) requirement set.</span></span>

> [!NOTE]
> <span data-ttu-id="165f7-117">O **conjunto de requisitos RibbonApi 1.1** ainda não tem suporte no manifesto, portanto, você não pode especificá-lo na seção do `<Requirements>` manifesto.</span><span class="sxs-lookup"><span data-stu-id="165f7-117">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span> <span data-ttu-id="165f7-118">Para testar o suporte, seu código deve chamar `Office.context.requirements.isSetSupported('RibbonApi', '1.1')` .</span><span class="sxs-lookup"><span data-stu-id="165f7-118">To test for support, your code should call `Office.context.requirements.isSetSupported('RibbonApi', '1.1')`.</span></span> <span data-ttu-id="165f7-119">Se e *somente se*, essa chamada retornar , seu código poderá chamar `true` as APIs habilitar/desabilitar.</span><span class="sxs-lookup"><span data-stu-id="165f7-119">If, *and only if*, that call returns `true`, your code can call the enable/disable APIs.</span></span> <span data-ttu-id="165f7-120">Se a chamada de retornar , todos os `isSetSupported` `false` comandos de complemento personalizados serão habilitados o tempo todo.</span><span class="sxs-lookup"><span data-stu-id="165f7-120">If the call of `isSetSupported` returns `false`, then all custom add-in commands are enabled all of the time.</span></span> <span data-ttu-id="165f7-121">Você deve projetar seu complemento de produção e qualquer instrução no aplicativo para levar em conta como ele funcionará quando o conjunto de requisitos **RibbonApi 1.1** não for suportado.</span><span class="sxs-lookup"><span data-stu-id="165f7-121">You must design your production add-in, and any in-app instructions, to take account of how it will work when the **RibbonApi 1.1** requirement set is not supported.</span></span> <span data-ttu-id="165f7-122">Para obter mais informações e exemplos de uso, consulte Especificar Office aplicativos e requisitos de API, especialmente Usar verificações de tempo de execução `isSetSupported` [em seu código JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code). [](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="165f7-122">For more information and examples of using `isSetSupported`, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md), especially [Use runtime checks in your JavaScript code](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="165f7-123">(A seção [Definir o elemento Requirements no manifesto](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) desse artigo não se aplica à Faixa de Opções 1.1.)</span><span class="sxs-lookup"><span data-stu-id="165f7-123">(The section [Set the Requirements element in the manifest](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest) of that article does not apply to Ribbon 1.1.)</span></span>

## <a name="shared-runtime-required"></a><span data-ttu-id="165f7-124">Tempo de execução compartilhado necessário</span><span class="sxs-lookup"><span data-stu-id="165f7-124">Shared runtime required</span></span>

<span data-ttu-id="165f7-125">As APIs e a marcação de manifesto descritas neste artigo exigem que o manifesto do suplemento especifique que ele deve usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="165f7-125">The APIs and manifest markup described in this article require that the add-in's manifest specify that it should use a shared runtime.</span></span> <span data-ttu-id="165f7-126">Para fazer isso, execute as seguintes etapas.</span><span class="sxs-lookup"><span data-stu-id="165f7-126">To do this take the following steps.</span></span>

1. <span data-ttu-id="165f7-127">No elemento [Runtimes](../reference/manifest/runtimes.md) no manifesto, adicione o seguinte elemento filho: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span><span class="sxs-lookup"><span data-stu-id="165f7-127">In the [Runtimes](../reference/manifest/runtimes.md) element in the manifest, add the following child element: `<Runtime resid="Contoso.SharedRuntime.Url" lifetime="long" />`.</span></span> <span data-ttu-id="165f7-128">(Se ainda não houver um elemento `<Runtimes>` no manifesto, crie-o como o primeiro filho abaixo do elemento `<Host>` na seção `VersionOverrides`.)</span><span class="sxs-lookup"><span data-stu-id="165f7-128">(If there isn't already a `<Runtimes>` element in the manifest, create it as the first child under the `<Host>` element in the `VersionOverrides` section.)</span></span>
2. <span data-ttu-id="165f7-129">Na seção [Resources.Urls](../reference/manifest/resources.md) do manifesto, adicione o seguinte elemento filho: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, onde `{MyDomain}` é o domínio do suplemento e `{path-to-start-page}` o caminho da página inicial do suplemento; por exemplo: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span><span class="sxs-lookup"><span data-stu-id="165f7-129">In the [Resources.Urls](../reference/manifest/resources.md) section of the manifest, add the following child element: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://{MyDomain}/{path-to-start-page}" />`, where `{MyDomain}` is the domain of the add-in and `{path-to-start-page}` is the path for the start page of the add-in; for example: `<bt:Url id="Contoso.SharedRuntime.Url" DefaultValue="https://localhost:3000/index.html" />`.</span></span>
3. <span data-ttu-id="165f7-130">Dependendo do seu suplemento conter um painel de tarefas, um arquivo de função ou uma função personalizada do Excel, você deve executar uma ou mais das três etapas a seguir:</span><span class="sxs-lookup"><span data-stu-id="165f7-130">Depending on whether your add-in contains a task pane, a function file, or an Excel custom function, you must do one or more of the following three steps:</span></span>

    - <span data-ttu-id="165f7-131">Se o suplemento contiver um painel de tarefas, defina o `resid` atributo do elemento [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do elemento `<Runtime>` na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="165f7-131">If the add-in contains a task pane, set the `resid` attribute of the [Action](../reference/manifest/action.md).[SourceLocation](../reference/manifest/sourcelocation.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="165f7-132">O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="165f7-132">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="165f7-133">Se o suplemento contiver uma função personalizada do Excel, defina o `resid` atributo do elemento [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) para exatamente a mesma série de caracteres que você usou para `resid` do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="165f7-133">If the add-in contains an Excel custom function, set the `resid` attribute of the [Page](../reference/manifest/page.md).[SourceLocation](../reference/manifest/sourcelocation.md) element exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="165f7-134">O elemento deve ficar assim: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="165f7-134">The element should look like this: `<SourceLocation resid="Contoso.SharedRuntime.Url"/>`.</span></span>
    - <span data-ttu-id="165f7-135">Se o suplemento contiver um arquivo de função, defina o `resid` atributo do elemento [FunctionFile](../reference/manifest/functionfile.md) para exatamente a mesma série que você usou para o `resid`do `<Runtime>` elemento na etapa 1. Por exemplo, `Contoso.SharedRuntime.Url`.</span><span class="sxs-lookup"><span data-stu-id="165f7-135">If the add-in contains a function file, set the `resid` attribute of the [FunctionFile](../reference/manifest/functionfile.md) element to exactly the same string as you used for the `resid` of the `<Runtime>` element in step 1; for example, `Contoso.SharedRuntime.Url`.</span></span> <span data-ttu-id="165f7-136">O elemento deve ficar assim: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span><span class="sxs-lookup"><span data-stu-id="165f7-136">The element should look like this: `<FunctionFile resid="Contoso.SharedRuntime.Url"/>`.</span></span>

## <a name="set-the-default-state-to-disabled"></a><span data-ttu-id="165f7-137">Defina o estado padrão como desabilitado</span><span class="sxs-lookup"><span data-stu-id="165f7-137">Set the default state to disabled</span></span>

<span data-ttu-id="165f7-138">Por padrão, qualquer comando de suplemento é habilitado quando o aplicativo do Office é iniciado.</span><span class="sxs-lookup"><span data-stu-id="165f7-138">By default, any Add-in Command is enabled when the Office application launches.</span></span> <span data-ttu-id="165f7-139">Se você deseja que um botão ou item de menu personalizado esteja desabilitado quando o aplicativo do Office for iniciado, especifique isso no manifesto.</span><span class="sxs-lookup"><span data-stu-id="165f7-139">If you want a custom button or menu item to be disabled when the Office application launches, you specify this in the manifest.</span></span> <span data-ttu-id="165f7-140">Basta adicionar um elemento [Enabled](../reference/manifest/enabled.md) (com o valor `false`) imediatamente *abaixo* (não dentro) do elemento [Ação](../reference/manifest/action.md) na declaração do controle.</span><span class="sxs-lookup"><span data-stu-id="165f7-140">Just add an [Enabled](../reference/manifest/enabled.md) element (with the value `false`) immediately *below* (not inside) the [Action](../reference/manifest/action.md) element in the declaration of the control.</span></span> <span data-ttu-id="165f7-141">Veja a estrutura básica a seguir:</span><span class="sxs-lookup"><span data-stu-id="165f7-141">The following shows the basic structure:</span></span>

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
                  ...
                  <Action ...>
                  <Enabled>false</Enabled>
...
</OfficeApp>
```

## <a name="change-the-state-programmatically"></a><span data-ttu-id="165f7-142">Alterar o estado programaticamente</span><span class="sxs-lookup"><span data-stu-id="165f7-142">Change the state programmatically</span></span>

<span data-ttu-id="165f7-143">As etapas essenciais para alterar o status habilitado de um Comando de Suplemento são:</span><span class="sxs-lookup"><span data-stu-id="165f7-143">The essential steps to changing the enabled status of an Add-in Command are:</span></span>

1. <span data-ttu-id="165f7-144">Crie um [objeto RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica o comando e seu grupo pai e guia, por suas IDs conforme declarado no manifesto; e (2) especifica o estado habilitado ou desabilitado do comando.</span><span class="sxs-lookup"><span data-stu-id="165f7-144">Create a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the command, and its parent group and tab, by their IDs as declared in the manifest; and (2) specifies the enabled or disabled state of the command.</span></span>
2. <span data-ttu-id="165f7-145">Passe o objeto **RibbonUpdaterData** para o método [OfficeRuntime.Ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-).</span><span class="sxs-lookup"><span data-stu-id="165f7-145">Pass the **RibbonUpdaterData** object to the [Office.ribbon.requestUpdate()](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method.</span></span>

<span data-ttu-id="165f7-146">Apresentamos um exemplo simples a seguir.</span><span class="sxs-lookup"><span data-stu-id="165f7-146">The following is a simple example.</span></span> <span data-ttu-id="165f7-147">Observe que "MyButton", "OfficeAddinTab1" e "CustomGroup111" são copiados do manifesto.</span><span class="sxs-lookup"><span data-stu-id="165f7-147">Note that "MyButton", "OfficeAddinTab1", and "CustomGroup111" are copied from the manifest.</span></span>

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

<span data-ttu-id="165f7-148">Também fornecemos várias interfaces (tipos) para facilitar a construção do objeto **RibbonUpdateData**.</span><span class="sxs-lookup"><span data-stu-id="165f7-148">We also provide several interfaces (types) to make it easier to construct the **RibbonUpdateData** object.</span></span> <span data-ttu-id="165f7-149">Veja a seguir o exemplo equivalente no TypeScript, que faz uso desses tipos.</span><span class="sxs-lookup"><span data-stu-id="165f7-149">The following is the equivalent example in TypeScript and it makes use of these types.</span></span>

```typescript
const enableButton = async () => {
    const button: Control = {id: "MyButton", enabled: true};
    const parentGroup: Group = {id: "CustomGroup111", controls: [button]};
    const parentTab: Tab = {id: "OfficeAddinTab1", groups: [parentGroup]};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="165f7-150">Você pode `await` chamar **requestUpdate()** se a função pai for assíncrona, mas observe que o aplicativo Office controla quando atualiza o estado da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="165f7-150">You can `await` the call of **requestUpdate()** if the parent function is asynchronous, but note that the Office application controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="165f7-151">O método **requestUpdate()** adiciona uma solicitação para atualização à fila de espera.</span><span class="sxs-lookup"><span data-stu-id="165f7-151">The **requestUpdate()** method queues a request to update.</span></span> <span data-ttu-id="165f7-152">O método resolverá o objeto promise assim que tiver enraizado a solicitação, não quando a faixa de opções realmente for atualizada.</span><span class="sxs-lookup"><span data-stu-id="165f7-152">The method will resolve the promise object as soon as it has queued the request, not when the ribbon actually updates.</span></span>

## <a name="change-the-state-in-response-to-an-event"></a><span data-ttu-id="165f7-153">Alterar o estado em resposta a um evento</span><span class="sxs-lookup"><span data-stu-id="165f7-153">Change the state in response to an event</span></span>

<span data-ttu-id="165f7-154">Um cenário comum em que o estado da faixa de opções deve mudar é quando um evento iniciado pelo usuário altera o contexto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="165f7-154">A common scenario in which the ribbon state should change is when a user-initiated event changes the add-in context.</span></span>

<span data-ttu-id="165f7-155">Considere um cenário em que um botão deve ser ativado quando e somente quando um gráfico é ativado.</span><span class="sxs-lookup"><span data-stu-id="165f7-155">Consider a scenario in which a button should be enabled when, and only when, a chart is activated.</span></span> <span data-ttu-id="165f7-156">A primeira etapa é definir o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto como `false`.</span><span class="sxs-lookup"><span data-stu-id="165f7-156">The first step is to set the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest to `false`.</span></span> <span data-ttu-id="165f7-157">Veja um exemplo acima.</span><span class="sxs-lookup"><span data-stu-id="165f7-157">See above for an example.</span></span>

<span data-ttu-id="165f7-158">Segundo, atribua manipuladores.</span><span class="sxs-lookup"><span data-stu-id="165f7-158">Second, assign handlers.</span></span> <span data-ttu-id="165f7-159">Isso geralmente é feito no método **Office.onReady**, como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos eventos **onActivated** e **onDeactivated** de todos os gráficos da planilha.</span><span class="sxs-lookup"><span data-stu-id="165f7-159">This is commonly done in the **Office.onReady** method as in the following example which assigns handlers (created in a later step) to the **onActivated** and **onDeactivated** events of all the charts in the worksheet.</span></span>

```javascript
Office.onReady(async () => {
    await Excel.run(context => {
        var charts = context.workbook.worksheets
            .getActiveWorksheet()
            .charts;
        charts.onActivated.add(enableChartFormat);
        charts.onDeactivated.add(disableChartFormat);
        return context.sync();
    });
});
```

<span data-ttu-id="165f7-160">Terceiro, defina o manipulador `enableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="165f7-160">Third, define the `enableChartFormat` handler.</span></span> <span data-ttu-id="165f7-161">A seguir, é apresentado um exemplo simples, mas consulte [Prática recomendada: Teste se há erros de status do controle](#best-practice-test-for-control-status-errors) abaixo para obter uma maneira mais robusta de alterar o status de um controle.</span><span class="sxs-lookup"><span data-stu-id="165f7-161">The following is a simple example, but see [Best practice: Test for control status errors](#best-practice-test-for-control-status-errors) below for a more robust way of changing a control's status.</span></span>

```javascript
function enableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: true
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);
}
```

<span data-ttu-id="165f7-162">Quarto, defina o manipulador `disableChartFormat`.</span><span class="sxs-lookup"><span data-stu-id="165f7-162">Fourth, define the `disableChartFormat` handler.</span></span> <span data-ttu-id="165f7-163">Seria idêntico a `enableChartFormat`, exceto que a propriedade **enabled** do objeto button seria configurada como `false`.</span><span class="sxs-lookup"><span data-stu-id="165f7-163">It would be identical to `enableChartFormat` except that the **enabled** property of the button object would be set to `false`.</span></span>

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="165f7-164">Visibilidade da guia de alternância e o status habilitado de um botão ao mesmo tempo</span><span class="sxs-lookup"><span data-stu-id="165f7-164">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="165f7-165">O **método requestUpdate** também é usado para alternar a visibilidade de uma guia contextual personalizada. Para obter detalhes sobre isso e o código de exemplo, consulte [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).</span><span class="sxs-lookup"><span data-stu-id="165f7-165">The **requestUpdate** method is also used to toggle the visibility of a custom contextual tab. For details about this and example code, see [Create custom contextual tabs in Office Add-ins](contextual-tabs.md#toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time).</span></span>

## <a name="best-practice-test-for-control-status-errors"></a><span data-ttu-id="165f7-166">Prática recomendada: Teste se há erros de status do controle</span><span class="sxs-lookup"><span data-stu-id="165f7-166">Best practice: Test for control status errors</span></span>

<span data-ttu-id="165f7-167">Em algumas circunstâncias, a faixa de opções não é redesenhada após `requestUpdate` ser chamado, portanto, o status clicável do controle não muda.</span><span class="sxs-lookup"><span data-stu-id="165f7-167">In some circumstances, the ribbon does not repaint after `requestUpdate` is called, so the control's clickable status does not change.</span></span> <span data-ttu-id="165f7-168">Por esse motivo, é uma prática recomendada para o suplemento acompanhar o status de seus controles.</span><span class="sxs-lookup"><span data-stu-id="165f7-168">For this reason it is a best practice for the add-in to keep track of the status of its controls.</span></span> <span data-ttu-id="165f7-169">O suplemento deve estar em conformidade com estas regras:</span><span class="sxs-lookup"><span data-stu-id="165f7-169">The add-in should conform to these rules:</span></span>

1. <span data-ttu-id="165f7-170">Sempre que `requestUpdate` é chamado, o código deve registrar o estado pretendido dos botões e itens de menu personalizados.</span><span class="sxs-lookup"><span data-stu-id="165f7-170">Whenever `requestUpdate` is called, the code should record the intended state of the custom buttons and menu items.</span></span>
2. <span data-ttu-id="165f7-171">Quando um controle personalizado é clicado, o primeiro código no manipulador deve verificar se o botão deveria ter sido clicável.</span><span class="sxs-lookup"><span data-stu-id="165f7-171">When a custom control is clicked, the first code in the handler, should check to see if the button should have been clickable.</span></span> <span data-ttu-id="165f7-172">Se não deveria ter sido, o código deve relatar ou registrar um erro e tentar novamente definir os botões no estado pretendido.</span><span class="sxs-lookup"><span data-stu-id="165f7-172">If shouldn't have been, the code should report or log an error and try again to set the buttons to the intended state.</span></span>

<span data-ttu-id="165f7-173">O exemplo a seguir mostra uma função que desativa um botão e registra o status do botão.</span><span class="sxs-lookup"><span data-stu-id="165f7-173">The following example shows a function that disables a button and records the button's status.</span></span> <span data-ttu-id="165f7-174">Observe que `chartFormatButtonEnabled` é uma variável booleana global inicializada com o mesmo valor que o elemento [Enabled](../reference/manifest/enabled.md) para o botão no manifesto.</span><span class="sxs-lookup"><span data-stu-id="165f7-174">Note that `chartFormatButtonEnabled` is a global boolean variable that is initialized to the same value as the [Enabled](../reference/manifest/enabled.md) element for the button in the manifest.</span></span>

```javascript
function disableChartFormat() {
    var button = {
                  id: "ChartFormatButton", 
                  enabled: false
                 };
    var parentGroup = {
                       id: "MyGroup",
                       controls: [button]
                      };
    var parentTab = {
                     id: "CustomChartTab", 
                     groups: [parentGroup]
                    };
    var ribbonUpdater = {tabs: [parentTab]};
    Office.ribbon.requestUpdate(ribbonUpdater);

    chartFormatButtonEnabled = false;
}
```

<span data-ttu-id="165f7-175">O exemplo a seguir mostra como o manipulador do botão testa um estado incorreto do botão.</span><span class="sxs-lookup"><span data-stu-id="165f7-175">The following example shows how the button's handler tests for an incorrect state of the button.</span></span> <span data-ttu-id="165f7-176">Observe que `reportError` é uma função que mostra ou registra um erro.</span><span class="sxs-lookup"><span data-stu-id="165f7-176">Note that `reportError` is a function that shows or logs an error.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="165f7-177">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="165f7-177">Error handling</span></span>

<span data-ttu-id="165f7-178">Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="165f7-178">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="165f7-179">Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="165f7-179">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="165f7-180">Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="165f7-180">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="165f7-181">Veja um exemplo de como lidar com esse erro a seguir.</span><span class="sxs-lookup"><span data-stu-id="165f7-181">The following is an example of how to handle this error.</span></span> <span data-ttu-id="165f7-182">Nesse caso, o método `reportError` exibe o erro para o usuário.</span><span class="sxs-lookup"><span data-stu-id="165f7-182">In this case, the `reportError` method displays the error to the user.</span></span>

```javascript
function disableChartFormat() {
    try {
        var button = {
                      id: "ChartFormatButton", 
                      enabled: false
                     };
        var parentGroup = {
                           id: "MyGroup",
                           controls: [button]
                          };
        var parentTab = {
                         id: "CustomChartTab", 
                         groups: [parentGroup]
                        };
        var ribbonUpdater = {tabs: [parentTab]};
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
