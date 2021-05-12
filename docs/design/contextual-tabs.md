---
title: Criar guias contextuais personalizadas em Office de complementos
description: Saiba como adicionar guias contextuais personalizadas ao seu Office Add-in.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 980beb24a3d384ecf21da44db288272a1ab1b0e3
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330168"
---
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="16853-103">Criar guias contextuais personalizadas em Office de complementos</span><span class="sxs-lookup"><span data-stu-id="16853-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="16853-104">Uma guia contextual é um controle de tabulação oculto na faixa Office que é exibido na linha de tabulação quando um evento especificado ocorre no documento Office.</span><span class="sxs-lookup"><span data-stu-id="16853-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="16853-105">Por exemplo, a **guia Design** de Tabela que aparece na faixa Excel quando uma tabela é selecionada.</span><span class="sxs-lookup"><span data-stu-id="16853-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="16853-106">Você pode incluir guias contextuais personalizadas no seu Office e especificar quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="16853-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="16853-107">(No entanto, as guias contextuais personalizadas não respondem a alterações de foco.)</span><span class="sxs-lookup"><span data-stu-id="16853-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="16853-108">Este artigo pressupõe que você esteja familiarizado com a seguinte documentação.</span><span class="sxs-lookup"><span data-stu-id="16853-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="16853-109">Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).</span><span class="sxs-lookup"><span data-stu-id="16853-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="16853-110">Conceitos básicos dos Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="16853-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="16853-111">No momento, as guias contextuais personalizadas só têm suporte no Excel e somente nessas plataformas e builds:</span><span class="sxs-lookup"><span data-stu-id="16853-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="16853-112">Excel no Windows (somente Microsoft 365 assinatura): Versão 2102 (Build 13801.20294) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="16853-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="16853-113">Excel Online</span><span class="sxs-lookup"><span data-stu-id="16853-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="16853-114">As guias contextuais personalizadas funcionam somente em plataformas que suportam os seguintes conjuntos de requisitos.</span><span class="sxs-lookup"><span data-stu-id="16853-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="16853-115">Para obter mais informações sobre conjuntos de requisitos e como trabalhar com eles, consulte [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="16853-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="16853-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16853-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="16853-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="16853-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="16853-118">Você pode usar as verificações de tempo de execução em seu código para testar se a combinação de host e plataforma do usuário oferece suporte a esses conjuntos de requisitos, conforme descrito em Especificar aplicativos Office e requisitos [de API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="16853-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="16853-119">(A técnica de especificar os conjuntos de requisitos no manifesto, que também é descrito nesse artigo, não funciona atualmente para RibbonApi 1.2.) Como alternativa, você pode [implementar uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span><span class="sxs-lookup"><span data-stu-id="16853-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="16853-120">Comportamento de guias contextuais personalizadas</span><span class="sxs-lookup"><span data-stu-id="16853-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="16853-121">A experiência do usuário para guias contextuais personalizadas segue o padrão de guias Office contextuais internas.</span><span class="sxs-lookup"><span data-stu-id="16853-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="16853-122">Veja a seguir os princípios básicos para as guias contextuais personalizadas de posicionamento:</span><span class="sxs-lookup"><span data-stu-id="16853-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="16853-123">Quando uma guia contextual personalizada é visível, ela aparece na extremidade direita da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="16853-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="16853-124">Se uma ou mais guias contextuais e uma ou mais guias contextuais personalizadas de complementos são visíveis ao mesmo tempo, as guias contextuais personalizadas estão sempre à direita de todas as guias contextuais.</span><span class="sxs-lookup"><span data-stu-id="16853-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="16853-125">Se o seu add-in tiver mais de uma guia contextual e houver contextos nos quais mais de um está visível, eles aparecerão na ordem em que são definidos no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="16853-126">(A direção é a mesma direção que o idioma Office, ou seja, é da esquerda para a direita em idiomas da esquerda para a direita, mas da direita para a esquerda em idiomas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como defini-los.</span><span class="sxs-lookup"><span data-stu-id="16853-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="16853-127">Se mais de um complemento tiver uma guia contextual visível em um contexto específico, elas aparecerão na ordem na qual os complementos foram lançados.</span><span class="sxs-lookup"><span data-stu-id="16853-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="16853-128">Guias *contextuais personalizadas,* ao contrário das guias principais personalizadas, não são adicionadas permanentemente à faixa Office do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="16853-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="16853-129">Eles estão presentes apenas em Office documentos nos quais o seu complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="16853-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="16853-130">Principais etapas para incluir uma guia contextual em um complemento</span><span class="sxs-lookup"><span data-stu-id="16853-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="16853-131">Veja a seguir as principais etapas para incluir uma guia contextual personalizada em um complemento:</span><span class="sxs-lookup"><span data-stu-id="16853-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="16853-132">Configure o complemento para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="16853-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="16853-133">Defina a guia e os grupos e controles que aparecem nele.</span><span class="sxs-lookup"><span data-stu-id="16853-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="16853-134">Registre a guia contextual com Office.</span><span class="sxs-lookup"><span data-stu-id="16853-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="16853-135">Especifique as circunstâncias em que a guia ficará visível.</span><span class="sxs-lookup"><span data-stu-id="16853-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="16853-136">Configurar o add-in para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="16853-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="16853-137">A adição de guias contextuais personalizadas exige que o seu add-in use o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="16853-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="16853-138">Para obter mais informações, [consulte Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="16853-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="16853-139">Definir os grupos e controles que aparecem na guia</span><span class="sxs-lookup"><span data-stu-id="16853-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="16853-140">Ao contrário das guias principais personalizadas, definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com um blob JSON.</span><span class="sxs-lookup"><span data-stu-id="16853-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="16853-141">Seu código analisará o blob em um objeto JavaScript e passará o objeto para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="16853-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="16853-142">As guias contextuais personalizadas só estão presentes em documentos nos quais o seu complemento está sendo executado no momento.</span><span class="sxs-lookup"><span data-stu-id="16853-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="16853-143">Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções do aplicativo Office quando o complemento é instalado e permanecem presentes quando outro documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="16853-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="16853-144">Além disso, `requestCreateControls` o método pode ser executado apenas uma vez em uma sessão do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="16853-145">Se for chamado novamente, será lançado um erro.</span><span class="sxs-lookup"><span data-stu-id="16853-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="16853-146">A estrutura das propriedades e subpropropriedades do blob JSON (e os nomes principais) é aproximadamente paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="16853-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="16853-147">Construiremos um exemplo de guias contextuais blob JSON passo a passo.</span><span class="sxs-lookup"><span data-stu-id="16853-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="16853-148">(O esquema completo da guia contextual JSON estádynamic-ribbon.schema.js[ em](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="16853-148">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="16853-149">Esse link pode não estar funcionando no período de visualização para guias contextuais.</span><span class="sxs-lookup"><span data-stu-id="16853-149">This link may not be working in the preview period for contextual tabs.</span></span> <span data-ttu-id="16853-150">Se o link não estiver funcionando, você poderá encontrar o rascunho mais recente do esquema no [rascunho dynamic-ribbon.schema.jsem](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) Se você estiver trabalhando no Visual Studio Code, poderá usar esse arquivo para obter IntelliSense e validar seu JSON.</span><span class="sxs-lookup"><span data-stu-id="16853-150">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon/1.0/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="16853-151">Para obter mais informações, [consulte Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="16853-151">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="16853-152">Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz nomeadas `actions` e `tabs` .</span><span class="sxs-lookup"><span data-stu-id="16853-152">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="16853-153">A `actions` matriz é uma especificação de todas as funções que podem ser executadas por controles na guia contextual. A matriz define uma ou mais guias `tabs` contextuais, *até um máximo de 20*.</span><span class="sxs-lookup"><span data-stu-id="16853-153">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="16853-154">Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação.</span><span class="sxs-lookup"><span data-stu-id="16853-154">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="16853-155">Adicione o seguinte como o único membro da `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-155">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="16853-156">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-156">About this markup, note:</span></span>

    - <span data-ttu-id="16853-157">As `id` propriedades e são `type` obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="16853-157">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="16853-158">O valor de `type` pode ser "ExecuteFunction" ou "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="16853-158">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="16853-159">A `functionName` propriedade só é usada quando o valor de é `type` `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="16853-159">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="16853-160">É o nome de uma função definida no FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="16853-160">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="16853-161">Para obter mais informações sobre o FunctionFile, consulte [Conceitos básicos para Comandos de Complemento.](add-in-commands.md)</span><span class="sxs-lookup"><span data-stu-id="16853-161">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="16853-162">Em uma etapa posterior, você mapeará essa ação para um botão na guia contextual.</span><span class="sxs-lookup"><span data-stu-id="16853-162">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="16853-163">Adicione o seguinte como o único membro da `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-163">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="16853-164">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-164">About this markup, note:</span></span>

    - <span data-ttu-id="16853-165">A propriedade `id` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="16853-165">The `id` property is required.</span></span> <span data-ttu-id="16853-166">Use uma ID breve e descritiva que seja exclusiva entre todas as guias contextuais no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-166">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="16853-167">A propriedade `label` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="16853-167">The `label` property is required.</span></span> <span data-ttu-id="16853-168">É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.</span><span class="sxs-lookup"><span data-stu-id="16853-168">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="16853-169">A propriedade `groups` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="16853-169">The `groups` property is required.</span></span> <span data-ttu-id="16853-170">Ele define os grupos de controles que serão exibidos na guia. Ele deve ter pelo menos um membro *e não mais de 20*.</span><span class="sxs-lookup"><span data-stu-id="16853-170">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="16853-171">(Também há limites sobre o número de controles que você pode ter em uma guia contextual personalizada e que também restringirá quantos grupos você tem.</span><span class="sxs-lookup"><span data-stu-id="16853-171">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="16853-172">Consulte a próxima etapa para obter mais informações.)</span><span class="sxs-lookup"><span data-stu-id="16853-172">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="16853-173">O objeto tab também pode ter uma propriedade opcional que especifica se a guia fica `visible` visível imediatamente quando o complemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="16853-173">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="16853-174">Como as guias contextuais normalmente ficam ocultas até que um evento do usuário acione sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a propriedade é padrão quando não está `visible` `false` presente.</span><span class="sxs-lookup"><span data-stu-id="16853-174">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="16853-175">Em uma seção posterior, mostramos como definir a propriedade como `true` em resposta a um evento.</span><span class="sxs-lookup"><span data-stu-id="16853-175">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="16853-176">No exemplo simples em andamento, a guia contextual tem apenas um único grupo.</span><span class="sxs-lookup"><span data-stu-id="16853-176">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="16853-177">Adicione o seguinte como o único membro da `groups` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-177">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="16853-178">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-178">About this markup, note:</span></span>

    - <span data-ttu-id="16853-179">Todas as propriedades são necessárias.</span><span class="sxs-lookup"><span data-stu-id="16853-179">All the properties are required.</span></span>
    - <span data-ttu-id="16853-180">A `id` propriedade deve ser exclusiva entre todos os grupos na guia. Use uma ID breve e descritiva.</span><span class="sxs-lookup"><span data-stu-id="16853-180">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="16853-181">A `label` é uma cadeia de caracteres amigável para servir como o rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="16853-181">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="16853-182">O valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela Office `icon` aplicativo.</span><span class="sxs-lookup"><span data-stu-id="16853-182">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="16853-183">O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e os menus no grupo.</span><span class="sxs-lookup"><span data-stu-id="16853-183">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="16853-184">Deve haver pelo menos um.</span><span class="sxs-lookup"><span data-stu-id="16853-184">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="16853-185">*O número total de controles na guia inteira não pode ser maior do que 20.*</span><span class="sxs-lookup"><span data-stu-id="16853-185">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="16853-186">Por exemplo, você pode ter 3 grupos com 6 controles cada e um quarto grupo com 2 controles, mas não pode ter 4 grupos com 6 controles cada.</span><span class="sxs-lookup"><span data-stu-id="16853-186">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="16853-187">Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32x32 px e 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="16853-187">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="16853-188">Opcionalmente, você também pode ter ícones de tamanhos 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="16853-188">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="16853-189">Office decide qual ícone usar com base no tamanho da faixa de opções e Office de aplicativos.</span><span class="sxs-lookup"><span data-stu-id="16853-189">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="16853-190">Adicione os seguintes objetos à matriz de ícones.</span><span class="sxs-lookup"><span data-stu-id="16853-190">Add the following objects to the icon array.</span></span> <span data-ttu-id="16853-191">(Se os tamanhos da janela e da  faixa de opções são grandes o suficiente para que pelo menos um dos controles no grupo apareça, nenhum ícone de grupo será exibido.</span><span class="sxs-lookup"><span data-stu-id="16853-191">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="16853-192">Por exemplo, assista ao grupo **Estilos** na faixa de opções do Word enquanto você reduz e expande a janela do Word.) Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-192">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="16853-193">Ambas as propriedades são necessárias.</span><span class="sxs-lookup"><span data-stu-id="16853-193">Both the properties are required.</span></span>
    - <span data-ttu-id="16853-194">A `size` unidade de medida da propriedade é pixels.</span><span class="sxs-lookup"><span data-stu-id="16853-194">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="16853-195">Os ícones são sempre quadrados, portanto, o número é a altura e a largura.</span><span class="sxs-lookup"><span data-stu-id="16853-195">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="16853-196">A `sourceLocation` propriedade especifica a URL completa para o ícone.</span><span class="sxs-lookup"><span data-stu-id="16853-196">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="16853-197">Assim como você normalmente deve alterar as URLs no manifesto do complemento quando você muda do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.</span><span class="sxs-lookup"><span data-stu-id="16853-197">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="16853-198">No nosso exemplo simples em andamento, o grupo tem apenas um único botão.</span><span class="sxs-lookup"><span data-stu-id="16853-198">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="16853-199">Adicione o objeto a seguir como o único membro da `controls` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-199">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="16853-200">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-200">About this markup, note:</span></span>

    - <span data-ttu-id="16853-201">Todas as propriedades, exceto `enabled` , são necessárias.</span><span class="sxs-lookup"><span data-stu-id="16853-201">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="16853-202">`type` especifica o tipo de controle.</span><span class="sxs-lookup"><span data-stu-id="16853-202">`type` specifies the type of control.</span></span> <span data-ttu-id="16853-203">Os valores podem ser "Button", "Menu" ou "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="16853-203">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="16853-204">`id` pode ter até 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="16853-204">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="16853-205">`actionId` deve ser a ID de uma ação definida na `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-205">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="16853-206">(Consulte a etapa 1 desta seção.)</span><span class="sxs-lookup"><span data-stu-id="16853-206">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="16853-207">`label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.</span><span class="sxs-lookup"><span data-stu-id="16853-207">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="16853-208">`superTip` representa uma forma rica de dica de ferramenta.</span><span class="sxs-lookup"><span data-stu-id="16853-208">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="16853-209">As propriedades `title` e `description` são necessárias.</span><span class="sxs-lookup"><span data-stu-id="16853-209">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="16853-210">`icon` especifica os ícones do botão.</span><span class="sxs-lookup"><span data-stu-id="16853-210">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="16853-211">Os comentários anteriores sobre o ícone de grupo também se aplicam aqui.</span><span class="sxs-lookup"><span data-stu-id="16853-211">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="16853-212">`enabled` (opcional) especifica se o botão está habilitado quando a guia contextual aparece é iniciada.</span><span class="sxs-lookup"><span data-stu-id="16853-212">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="16853-213">O padrão se não estiver presente é `true` .</span><span class="sxs-lookup"><span data-stu-id="16853-213">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="16853-214">Veja a seguir o exemplo completo do blob JSON:</span><span class="sxs-lookup"><span data-stu-id="16853-214">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="16853-215">Registrar a guia contextual com Office requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="16853-215">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="16853-216">A guia contextual é registrada com Office chamando o [método Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="16853-216">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="16853-217">Isso normalmente é feito na função atribuída ou `Office.initialize` com o `Office.onReady` método.</span><span class="sxs-lookup"><span data-stu-id="16853-217">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="16853-218">Para saber mais sobre esses métodos e inicializar o add-in, consulte [Initialize your Office Add-in](../develop/initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="16853-218">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="16853-219">No entanto, você pode chamar o método a qualquer momento após a inicialização.</span><span class="sxs-lookup"><span data-stu-id="16853-219">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="16853-220">O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-220">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="16853-221">Um erro será lançado se for chamado novamente.</span><span class="sxs-lookup"><span data-stu-id="16853-221">An error is thrown if it is called again.</span></span>

<span data-ttu-id="16853-222">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="16853-222">The following is an example.</span></span> <span data-ttu-id="16853-223">Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o método antes de poder ser passada para `JSON.parse` uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="16853-223">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="16853-224">Especifique os contextos quando a guia ficará visível com requestUpdate</span><span class="sxs-lookup"><span data-stu-id="16853-224">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="16853-225">Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto do complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-225">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="16853-226">Considere um cenário no qual a guia deve estar visível quando, e somente quando, um gráfico (na planilha padrão de uma pasta de trabalho Excel) é ativado.</span><span class="sxs-lookup"><span data-stu-id="16853-226">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="16853-227">Comece atribuindo manipuladores.</span><span class="sxs-lookup"><span data-stu-id="16853-227">Begin by assigning handlers.</span></span> <span data-ttu-id="16853-228">Isso é geralmente feito no método como no exemplo a seguir que atribui manipuladores (criados em uma etapa posterior) aos eventos e de todos os `Office.onReady` `onActivated` `onDeactivated` gráficos na planilha.</span><span class="sxs-lookup"><span data-stu-id="16853-228">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="16853-229">Em seguida, defina os manipuladores.</span><span class="sxs-lookup"><span data-stu-id="16853-229">Next, define the handlers.</span></span> <span data-ttu-id="16853-230">Veja a seguir um exemplo simples de um , mas consulte Manipulando o `showDataTab` [erro HostRestartNeeded](#handle-the-hostrestartneeded-error) posteriormente neste artigo para obter uma versão mais robusta da função.</span><span class="sxs-lookup"><span data-stu-id="16853-230">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="16853-231">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="16853-231">About this code, note:</span></span>

- <span data-ttu-id="16853-232">O Office controla quando atualiza o estado da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="16853-232">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="16853-233">O [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) enfileia uma solicitação para atualizar.</span><span class="sxs-lookup"><span data-stu-id="16853-233">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="16853-234">O método resolverá o objeto assim que tiver enraizado a solicitação, não quando a faixa `Promise` de opções realmente for atualizada.</span><span class="sxs-lookup"><span data-stu-id="16853-234">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="16853-235">O parâmetro para o método é um objeto `requestUpdate` [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia por sua ID exatamente como especificado no *JSON* e (2) especifica a visibilidade da guia.</span><span class="sxs-lookup"><span data-stu-id="16853-235">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="16853-236">Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar objetos de tabulação adicionais à `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="16853-236">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="16853-237">O manipulador para ocultar a guia é quase idêntico, exceto pelo fato de que ela define a `visible` propriedade de volta como `false` .</span><span class="sxs-lookup"><span data-stu-id="16853-237">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="16853-238">A Office javaScript também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto.</span><span class="sxs-lookup"><span data-stu-id="16853-238">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="16853-239">A seguir está `showDataTab` a função em TypeScript e faz uso desses tipos.</span><span class="sxs-lookup"><span data-stu-id="16853-239">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="16853-240">Visibilidade da guia de alternância e o status habilitado de um botão ao mesmo tempo</span><span class="sxs-lookup"><span data-stu-id="16853-240">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="16853-241">O método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada `requestUpdate` ou em uma guia principal personalizada. Para obter detalhes sobre isso, consulte [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="16853-241">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="16853-242">Pode haver cenários nos quais você deseja alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="16853-242">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="16853-243">Você pode fazer isso com uma única chamada de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="16853-243">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="16853-244">A seguir está um exemplo no qual um botão em uma guia principal é habilitado ao mesmo tempo em que uma guia contextual é visível.</span><span class="sxs-lookup"><span data-stu-id="16853-244">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="16853-245">No exemplo a seguir, o botão habilitado está na mesma guia contextual que está sendo visível.</span><span class="sxs-lookup"><span data-stu-id="16853-245">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="16853-246">Localizando o blob JSON</span><span class="sxs-lookup"><span data-stu-id="16853-246">Localizing the JSON blob</span></span>

<span data-ttu-id="16853-247">O blob JSON que é passado não é localizado da mesma maneira que a marcação de manifesto para guias principais personalizadas é localizada (que é descrita em Localização de Controle do `requestCreateControls` [manifesto](../develop/localization.md#control-localization-from-the-manifest)).</span><span class="sxs-lookup"><span data-stu-id="16853-247">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="16853-248">Em vez disso, a localização deve ocorrer em tempo de execução usando blobs JSON distintos para cada localidade.</span><span class="sxs-lookup"><span data-stu-id="16853-248">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="16853-249">Sugerimos que você use uma instrução que testa a `switch` [propriedade Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="16853-249">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="16853-250">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="16853-250">The following is an example:</span></span>

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

<span data-ttu-id="16853-251">Em seguida, seu código chama a função para obter o blob localizado que é passado para `requestCreateControls` , como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="16853-251">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="16853-252">Práticas recomendadas para guias contextuais personalizadas</span><span class="sxs-lookup"><span data-stu-id="16853-252">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="16853-253">Implementar uma experiência de interface do usuário alternativa quando guias contextuais personalizadas não são suportadas</span><span class="sxs-lookup"><span data-stu-id="16853-253">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="16853-254">Algumas combinações de plataforma, Office aplicativo e Office build não suportam `requestCreateControls` .</span><span class="sxs-lookup"><span data-stu-id="16853-254">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="16853-255">Seu complemento deve ser projetado para fornecer uma experiência alternativa aos usuários que estão executando o complemento em uma dessas combinações.</span><span class="sxs-lookup"><span data-stu-id="16853-255">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="16853-256">As seções a seguir descrevem duas maneiras de fornecer uma experiência de fallback.</span><span class="sxs-lookup"><span data-stu-id="16853-256">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="16853-257">Usar guias ou controles nãocontextuais</span><span class="sxs-lookup"><span data-stu-id="16853-257">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="16853-258">Há um elemento de manifesto, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), projetado para criar uma experiência de fallback em um complemento que implementa guias contextuais personalizadas quando o add-in está sendo executado em um aplicativo ou plataforma que não oferece suporte a guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="16853-258">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="16853-259">A estratégia mais simples para usar esse elemento é que você define no manifesto uma ou mais guias principais personalizadas (ou seja, guias personalizadas *nãocontextuais)* que duplicam as personalizações da faixa de opções das guias contextuais personalizadas no seu complemento.</span><span class="sxs-lookup"><span data-stu-id="16853-259">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="16853-260">Mas você adiciona `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` como o primeiro elemento filho do [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="16853-260">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="16853-261">O efeito de fazer isso é o seguinte:</span><span class="sxs-lookup"><span data-stu-id="16853-261">The effect of doing so is the following:</span></span>

- <span data-ttu-id="16853-262">Se o complemento for executado em um aplicativo e plataforma que suportam guias contextuais personalizadas, a guia principal personalizada não aparecerá na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="16853-262">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="16853-263">Em vez disso, a guia contextual personalizada será criada quando o complemento chamar o `requestCreateControls` método.</span><span class="sxs-lookup"><span data-stu-id="16853-263">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="16853-264">Se o complemento for executado em  um aplicativo ou plataforma que não oferece suporte, a guia principal personalizada `requestCreateControls` aparecerá na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="16853-264">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="16853-265">A seguir, um exemplo dessa estratégia simples.</span><span class="sxs-lookup"><span data-stu-id="16853-265">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="16853-266">Essa estratégia simples usa uma guia principal personalizada que espelha uma guia contextual personalizada com seus grupos e controles filho, mas você pode usar uma estratégia mais complexa.</span><span class="sxs-lookup"><span data-stu-id="16853-266">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="16853-267">O elemento também pode ser adicionado como (o primeiro) elemento filho aos elementos Group e Control (tipo de botão e tipo de menu ) e `<OverriddenByRibbonApi>` elementos de [](../reference/manifest/control.md) [](../reference/manifest/group.md) [](../reference/manifest/control.md#button-control) [](../reference/manifest/control.md#menu-dropdown-button-controls) `<Item>` menu.</span><span class="sxs-lookup"><span data-stu-id="16853-267">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="16853-268">Esse fato permite distribuir os grupos e controles que, de outra forma, apareceriam na guia contextual entre vários grupos, botões e menus em várias guias principais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="16853-268">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="16853-269">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="16853-269">The following is an example.</span></span> <span data-ttu-id="16853-270">Observe que "MyButton" aparecerá na guia principal personalizada somente quando as guias contextuais personalizadas não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="16853-270">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="16853-271">Mas o grupo pai e a guia principal personalizada aparecerão independentemente de as guias contextuais personalizadas são suportadas.</span><span class="sxs-lookup"><span data-stu-id="16853-271">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="16853-272">Para obter mais exemplos, consulte [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="16853-272">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="16853-273">Quando uma guia pai, grupo ou menu é marcado com , ela não fica visível e toda a marcação filha é ignorada, quando as guias contextuais personalizadas não são `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` suportadas.</span><span class="sxs-lookup"><span data-stu-id="16853-273">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="16853-274">Portanto, não importa se qualquer um desses elementos filho tem o `<OverriddenByRibbonApi>` elemento ou qual é seu valor.</span><span class="sxs-lookup"><span data-stu-id="16853-274">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="16853-275">A implicação disso é que, se um item de menu, controle ou grupo deve estar visível em todos os contextos, então não só não deve ser marcado com , mas seu menu ancestral, grupo e guia também não deve ser marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` *dessa forma*.</span><span class="sxs-lookup"><span data-stu-id="16853-275">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="16853-276">Não marque todos *os* elementos filho de uma guia, grupo ou menu com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="16853-276">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="16853-277">Isso não faz sentido se o elemento pai for marcado com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` por motivos dados no parágrafo anterior.</span><span class="sxs-lookup"><span data-stu-id="16853-277">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="16853-278">Além disso, se você deixar de fora o no pai (ou defini-lo como ), o pai aparecerá independentemente de as guias contextuais personalizadas serem suportadas, mas elas estarão vazias quando elas são `<OverriddenByRibbonApi>` `false` suportadas.</span><span class="sxs-lookup"><span data-stu-id="16853-278">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="16853-279">Portanto, se todos os elementos filho não aparecerem quando as guias contextuais personalizadas são suportadas, marque o pai e somente o pai, com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="16853-279">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="16853-280">Usar APIs que mostram ou ocultam um painel de tarefas em contextos especificados</span><span class="sxs-lookup"><span data-stu-id="16853-280">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="16853-281">Como alternativa a , seu complemento pode definir um painel de tarefas com controles de interface do usuário que duplicam a funcionalidade dos controles em uma `<OverriddenByRibbonApi>` guia contextual personalizada. Em seguida, use os métodos [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) [e Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) para mostrar o painel de tarefas quando e somente quando a guia contextual teria sido mostrada se tivesse suporte.</span><span class="sxs-lookup"><span data-stu-id="16853-281">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="16853-282">Para obter detalhes sobre como usar esses métodos, consulte [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="16853-282">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="16853-283">Manipular o erro HostRestartNeeded</span><span class="sxs-lookup"><span data-stu-id="16853-283">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="16853-284">Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="16853-284">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="16853-285">Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="16853-285">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="16853-286">Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="16853-286">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="16853-287">Seu código deve lidar com esse erro.</span><span class="sxs-lookup"><span data-stu-id="16853-287">Your code should handle this error.</span></span> <span data-ttu-id="16853-288">Veja a seguir um exemplo de como.</span><span class="sxs-lookup"><span data-stu-id="16853-288">The following is an example of how.</span></span> <span data-ttu-id="16853-289">Nesse caso, o método `reportError` exibe o erro para o usuário.</span><span class="sxs-lookup"><span data-stu-id="16853-289">In this case, the `reportError` method displays the error to the user.</span></span>

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
