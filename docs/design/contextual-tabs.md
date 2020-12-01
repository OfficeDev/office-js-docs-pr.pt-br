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
# <a name="create-custom-contextual-tabs-in-office-add-ins-preview"></a><span data-ttu-id="fae55-103">Criar guias contextuais personalizadas em suplementos do Office (visualização)</span><span class="sxs-lookup"><span data-stu-id="fae55-103">Create custom contextual tabs in Office Add-ins (preview)</span></span>

<span data-ttu-id="fae55-104">Uma guia contextual é um controle de guia oculto na faixa de opções do Office que é exibido na linha da guia quando um evento especificado ocorre no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="fae55-105">Por exemplo, a guia **design da tabela** que aparece na faixa de opções do Excel quando uma tabela é selecionada.</span><span class="sxs-lookup"><span data-stu-id="fae55-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="fae55-106">Você pode incluir guias contextuais personalizadas no suplemento do Office e especificar quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="fae55-106">You can include custom contextual tabs in your Office add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="fae55-107">(No entanto, as guias contextuais personalizadas não respondem às alterações de foco.)</span><span class="sxs-lookup"><span data-stu-id="fae55-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="fae55-108">Este artigo pressupõe que você esteja familiarizado com a seguinte documentação.</span><span class="sxs-lookup"><span data-stu-id="fae55-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="fae55-109">Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).</span><span class="sxs-lookup"><span data-stu-id="fae55-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="fae55-110">Conceitos básicos dos Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="fae55-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="fae55-111">Guias contextuais personalizadas estão em versão prévia.</span><span class="sxs-lookup"><span data-stu-id="fae55-111">Custom contextual tabs are in preview.</span></span> <span data-ttu-id="fae55-112">Faça experiências com eles em um ambiente de desenvolvimento ou teste, mas não os adicione a um suplemento de produção.</span><span class="sxs-lookup"><span data-stu-id="fae55-112">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>
>
> <span data-ttu-id="fae55-113">Atualmente, as guias contextuais personalizadas só têm suporte no Excel e apenas nessas plataformas e compilações:</span><span class="sxs-lookup"><span data-stu-id="fae55-113">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="fae55-114">Excel no Windows (somente Microsoft 365, licença não permanente): versão 2011 (Build 13426,20274).</span><span class="sxs-lookup"><span data-stu-id="fae55-114">Excel on Windows (Microsoft 365 only, not perpetual license): Version 2011 (Build 13426.20274).</span></span> <span data-ttu-id="fae55-115">Sua assinatura do Microsoft 365 pode precisar estar no [canal atual (visualização)](https://insider.office.com/join/windows) , anteriormente chamado de "canal mensal (direcionado)" ou "insider Slow".</span><span class="sxs-lookup"><span data-stu-id="fae55-115">Your Microsoft 365 subscription may need to be on the [Current Channel (Preview)](https://insider.office.com/join/windows) formerly called "Monthly Channel (Targeted)" or "Insider Slow".</span></span>

> [!NOTE]
> <span data-ttu-id="fae55-116">Guias contextuais personalizadas funcionam somente em plataformas que dão suporte aos seguintes conjuntos de requisitos.</span><span class="sxs-lookup"><span data-stu-id="fae55-116">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="fae55-117">Para saber mais sobre conjuntos de requisitos e como trabalhar com eles, confira [especificar aplicativos do Office e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="fae55-117">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="fae55-118">SharedRuntime 1,1</span><span class="sxs-lookup"><span data-stu-id="fae55-118">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="fae55-119">Comportamento de guias contextuais personalizadas</span><span class="sxs-lookup"><span data-stu-id="fae55-119">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="fae55-120">A experiência do usuário para guias contextuais personalizadas segue o padrão de guias contextuais internas do Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-120">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="fae55-121">Estes são os princípios básicos para as guias contextuais personalizadas de posicionamento:</span><span class="sxs-lookup"><span data-stu-id="fae55-121">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="fae55-122">Quando uma guia contextual personalizada estiver visível, ela aparecerá na extremidade direita da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="fae55-122">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="fae55-123">Se uma ou mais guias contextuais internas e uma ou mais guias contextuais personalizadas de suplementos forem visíveis ao mesmo tempo, as guias contextuais personalizadas estarão sempre à direita de todas as guias contextuais internas.</span><span class="sxs-lookup"><span data-stu-id="fae55-123">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="fae55-124">Se o suplemento tiver mais de uma guia contextual e houver contextos em que mais de uma esteja visível, elas aparecerão na ordem em que estão definidas no suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae55-124">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="fae55-125">(A direção é a mesma direção do idioma do Office; ou seja, da esquerda para a direita em idiomas da esquerda para a direita, mas da direita para a esquerda em idiomas da direita para a esquerda.) Consulte [definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como você os define.</span><span class="sxs-lookup"><span data-stu-id="fae55-125">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="fae55-126">Se mais de um suplemento tiver uma guia contextual que seja visível em um contexto específico, elas aparecerão na ordem em que os suplementos foram iniciados.</span><span class="sxs-lookup"><span data-stu-id="fae55-126">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="fae55-127">As guias *contextuais* personalizadas, diferente das guias principais personalizadas, não são adicionadas permanentemente à faixa de opções do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-127">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="fae55-128">Eles estão presentes somente nos documentos do Office em que o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="fae55-128">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="fae55-129">Etapas principais para incluir uma guia contextual em um suplemento</span><span class="sxs-lookup"><span data-stu-id="fae55-129">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="fae55-130">A seguir estão as principais etapas para incluir uma guia contextual personalizada em um suplemento:</span><span class="sxs-lookup"><span data-stu-id="fae55-130">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="fae55-131">Configure o suplemento para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="fae55-131">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="fae55-132">Defina a guia e os grupos e controles que aparecem nele.</span><span class="sxs-lookup"><span data-stu-id="fae55-132">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="fae55-133">Registre a guia contextual com o Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-133">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="fae55-134">Especifique as circunstâncias em que a guia estará visível.</span><span class="sxs-lookup"><span data-stu-id="fae55-134">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="fae55-135">Configurar o suplemento para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="fae55-135">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="fae55-136">A adição de guias contextuais personalizadas exige que seu suplemento use o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="fae55-136">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="fae55-137">Para obter mais informações, consulte [configurar um suplemento para usar um tempo de execução compartilhado](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="fae55-137">For more information, see [Configure an add-in to use a shared runtime](../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="fae55-138">Definir os grupos e controles que aparecem na guia</span><span class="sxs-lookup"><span data-stu-id="fae55-138">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="fae55-139">Ao contrário das guias principais personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com um blob JSON.</span><span class="sxs-lookup"><span data-stu-id="fae55-139">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="fae55-140">O código analisa o blob em um objeto JavaScript e, em seguida, passa o objeto para o método [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) .</span><span class="sxs-lookup"><span data-stu-id="fae55-140">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="fae55-141">As guias contextuais personalizadas só estão presentes em documentos nos quais o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="fae55-141">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="fae55-142">Isso é diferente das guias principais personalizadas que são adicionadas à faixa de opções do aplicativo do Office quando o suplemento é instalado e permanecer presente quando outro documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="fae55-142">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="fae55-143">Além disso, o `requestCreateControls` método pode ser executado apenas uma vez em uma sessão do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae55-143">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="fae55-144">Se for chamado novamente, um erro será gerado.</span><span class="sxs-lookup"><span data-stu-id="fae55-144">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="fae55-145">A estrutura das propriedades e subpropriedades do blob JSON (e os nomes das chaves) é quase paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="fae55-145">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="fae55-146">Criaremos um exemplo de um blob JSON de guias contextual passo a passo.</span><span class="sxs-lookup"><span data-stu-id="fae55-146">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="fae55-147">(O esquema completo para a guia contextual JSON está em [dynamic-ribbon.schema.js](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span><span class="sxs-lookup"><span data-stu-id="fae55-147">(The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="fae55-148">Este link pode não estar funcionando no período de visualização inicial para guias contextuais.</span><span class="sxs-lookup"><span data-stu-id="fae55-148">This link may not be working in the early preview period for contextual tabs.</span></span> <span data-ttu-id="fae55-149">Se o link não estiver funcionando, você poderá encontrar o rascunho mais recente do esquema em [rascunho dynamic-ribbon.schema.jsem](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) Se você estiver trabalhando no Visual Studio Code, você pode usar esse arquivo para obter o IntelliSense e para validar seu JSON.</span><span class="sxs-lookup"><span data-stu-id="fae55-149">If the link is not working, you can find the latest draft of the schema at [draft dynamic-ribbon.schema.json](https://github.com/OfficeDev/testing-assets/tree/master/jsonschema/dynamic-ribbon.schema.json).) If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="fae55-150">Para obter mais informações, consulte [Editing JSON with Visual Studio Code-JSON schemas and Settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="fae55-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="fae55-151">Comece criando uma cadeia de caracteres JSON com duas propriedades de matriz chamadas `actions` e `tabs` .</span><span class="sxs-lookup"><span data-stu-id="fae55-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="fae55-152">A `actions` matriz é uma especificação de todas as funções que podem ser executadas pelos controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, *até o máximo de 10*.</span><span class="sxs-lookup"><span data-stu-id="fae55-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 10*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="fae55-153">Este exemplo simples de uma guia contextual terá apenas um único botão e, portanto, uma única ação.</span><span class="sxs-lookup"><span data-stu-id="fae55-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="fae55-154">Adicione o seguinte como o único membro da `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="fae55-155">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-155">About this markup, note:</span></span>

    - <span data-ttu-id="fae55-156">As `id` `type` Propriedades e são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="fae55-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="fae55-157">O valor de `type` pode ser "ExecuteFunction" ou "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="fae55-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="fae55-158">A `functionName` propriedade é usada somente quando o valor de `type` é `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="fae55-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="fae55-159">É o nome de uma função definida no Functionfile.</span><span class="sxs-lookup"><span data-stu-id="fae55-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="fae55-160">Para obter mais informações sobre o Functionfile, consulte [conceitos básicos para comandos de suplemento](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="fae55-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="fae55-161">Em uma etapa posterior, você irá mapear essa ação para um botão na guia contextual.</span><span class="sxs-lookup"><span data-stu-id="fae55-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="fae55-162">Adicione o seguinte como o único membro da `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="fae55-163">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-163">About this markup, note:</span></span>

    - <span data-ttu-id="fae55-164">A propriedade `id` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="fae55-164">The `id` property is required.</span></span> <span data-ttu-id="fae55-165">Use uma ID curta e descritiva exclusiva entre todas as guias contextuais no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae55-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="fae55-166">A propriedade `label` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="fae55-166">The `label` property is required.</span></span> <span data-ttu-id="fae55-167">É uma cadeia de caracteres amigável para servir como o rótulo da guia contextual.</span><span class="sxs-lookup"><span data-stu-id="fae55-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="fae55-168">A propriedade `groups` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="fae55-168">The `groups` property is required.</span></span> <span data-ttu-id="fae55-169">Ele define os grupos de controles que serão exibidos na guia. Deve ter pelo menos um membro e, no máximo *, 20*.</span><span class="sxs-lookup"><span data-stu-id="fae55-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="fae55-170">(Também há limites quanto ao número de controles que você pode ter em uma guia contextual personalizada e que também restringe o número de grupos que você tem.</span><span class="sxs-lookup"><span data-stu-id="fae55-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="fae55-171">Consulte a próxima etapa para obter mais informações.)</span><span class="sxs-lookup"><span data-stu-id="fae55-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="fae55-172">O objeto Tab também pode ter uma `visible` propriedade opcional que especifica se a guia estará visível imediatamente quando o suplemento for iniciado.</span><span class="sxs-lookup"><span data-stu-id="fae55-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="fae55-173">Como as guias contextuais são normalmente ocultas até que um evento de usuário dispare sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), a `visible` propriedade será definida como padrão `false` quando não estiver presente.</span><span class="sxs-lookup"><span data-stu-id="fae55-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="fae55-174">Em uma seção posterior, mostraremos como definir a propriedade como `true` em resposta a um evento.</span><span class="sxs-lookup"><span data-stu-id="fae55-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="fae55-175">No exemplo simples em andamento, a guia contextual tem apenas um único grupo.</span><span class="sxs-lookup"><span data-stu-id="fae55-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="fae55-176">Adicione o seguinte como o único membro da `groups` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="fae55-177">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-177">About this markup, note:</span></span>

    - <span data-ttu-id="fae55-178">Todas as propriedades são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="fae55-178">All the properties are required.</span></span>
    - <span data-ttu-id="fae55-179">A `id` propriedade deve ser exclusiva entre todos os grupos na guia. Use uma ID breve e descritiva.</span><span class="sxs-lookup"><span data-stu-id="fae55-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="fae55-180">O `label` é uma cadeia de caracteres amigável para servir como o rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="fae55-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="fae55-181">O `icon` valor da propriedade é uma matriz de objetos que especifica os ícones que o grupo terá na faixa de opções, dependendo do tamanho da faixa de opções e da janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="fae55-182">O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e outros controles no grupo.</span><span class="sxs-lookup"><span data-stu-id="fae55-182">The `controls` property's value is an array of objects that specify the buttons and other controls in the group.</span></span> <span data-ttu-id="fae55-183">Deve haver pelo menos um e *não mais de 6 em um grupo*.</span><span class="sxs-lookup"><span data-stu-id="fae55-183">There must be at least one and *no more than 6 in a group*.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="fae55-184">*O número total de controles na guia inteira não pode ser superior a 20.*</span><span class="sxs-lookup"><span data-stu-id="fae55-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="fae55-185">Por exemplo, você poderia ter 3 grupos com 6 controles cada e um quarto grupo com 2 controles, mas não pode ter quatro grupos com 6 controles cada.</span><span class="sxs-lookup"><span data-stu-id="fae55-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="fae55-186">Todos os grupos devem ter um ícone de pelo menos dois tamanhos, 32x32 PX e 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="fae55-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="fae55-187">Opcionalmente, você também pode ter ícones de tamanhos 16x16, 20x20, 24x24, 40x40, 48x48 e 64x64.</span><span class="sxs-lookup"><span data-stu-id="fae55-187">Optionally, you can also have icons of sizes 16x16, 20x20, 24x24, 40x40, 48x48 and 64x64.</span></span> <span data-ttu-id="fae55-188">O Office decide qual ícone usar com base no tamanho da faixa de opções e na janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="fae55-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="fae55-189">Adicione os seguintes objetos à matriz de ícones.</span><span class="sxs-lookup"><span data-stu-id="fae55-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="fae55-190">(Se a janela e os tamanhos de faixa de opções forem grandes o suficiente para que pelo menos um dos *controles* no grupo apareça, então nenhum ícone de grupo aparecerá.</span><span class="sxs-lookup"><span data-stu-id="fae55-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="fae55-191">Por exemplo, Assista ao grupo **estilos** na faixa de opções do Word à medida que você encolhe e expande a janela do Word.) Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="fae55-192">Ambas as propriedades são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="fae55-192">Both the properties are required.</span></span>
    - <span data-ttu-id="fae55-193">A `size` unidade de medida de propriedade é pixels.</span><span class="sxs-lookup"><span data-stu-id="fae55-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="fae55-194">Os ícones são sempre quadrados, portanto, o número é a altura e a largura.</span><span class="sxs-lookup"><span data-stu-id="fae55-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="fae55-195">A `sourceLocation` propriedade especifica a URL completa para o ícone.</span><span class="sxs-lookup"><span data-stu-id="fae55-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="fae55-196">Assim como você deve alterar as URLs no manifesto do suplemento quando migrar do desenvolvimento para a produção (como alterar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.</span><span class="sxs-lookup"><span data-stu-id="fae55-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="fae55-197">Em nosso exemplo simples em andamento, o grupo tem apenas um único botão.</span><span class="sxs-lookup"><span data-stu-id="fae55-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="fae55-198">Adicione o seguinte objeto como o único membro da `controls` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="fae55-199">Sobre essa marcação, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-199">About this markup, note:</span></span>

    - <span data-ttu-id="fae55-200">Todas as propriedades, exceto `enabled` , são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="fae55-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="fae55-201">`type` Especifica o tipo de controle.</span><span class="sxs-lookup"><span data-stu-id="fae55-201">`type` specifies the type of control.</span></span> <span data-ttu-id="fae55-202">Os valores podem ser "Button", "menu" ou "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="fae55-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="fae55-203">`id` pode ter até 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fae55-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="fae55-204">`actionId` deve ser a ID de uma ação definida na `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="fae55-205">(Confira a etapa 1 desta seção.)</span><span class="sxs-lookup"><span data-stu-id="fae55-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="fae55-206">`label` é uma cadeia de caracteres amigável para servir como o rótulo do botão.</span><span class="sxs-lookup"><span data-stu-id="fae55-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="fae55-207">`superTip` representa uma forma rica da dica de ferramenta.</span><span class="sxs-lookup"><span data-stu-id="fae55-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="fae55-208">As `title` Propriedades e `description` são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="fae55-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="fae55-209">`icon` Especifica os ícones para o botão.</span><span class="sxs-lookup"><span data-stu-id="fae55-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="fae55-210">Os comentários anteriores sobre o ícone de grupo aplicam-se aqui também.</span><span class="sxs-lookup"><span data-stu-id="fae55-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="fae55-211">`enabled` (opcional) especifica se o botão está habilitado quando a guia contextual aparece é iniciada.</span><span class="sxs-lookup"><span data-stu-id="fae55-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="fae55-212">O padrão, se não estiver presente, é `true` .</span><span class="sxs-lookup"><span data-stu-id="fae55-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="fae55-213">Veja a seguir o exemplo completo do blob JSON:</span><span class="sxs-lookup"><span data-stu-id="fae55-213">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="fae55-214">Registrar a guia contextual com o Office com o requestCreateControls</span><span class="sxs-lookup"><span data-stu-id="fae55-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="fae55-215">A guia contextual é registrada no Office chamando o método [Office. Ribbon. requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) .</span><span class="sxs-lookup"><span data-stu-id="fae55-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="fae55-216">Isso geralmente é feito na função que é atribuída `Office.initialize` ou ao `Office.onReady` método.</span><span class="sxs-lookup"><span data-stu-id="fae55-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="fae55-217">Para saber mais sobre esses métodos e inicializar o suplemento, confira [inicializar o suplemento do Office](../develop/initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="fae55-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="fae55-218">No entanto, você pode chamar o método a qualquer momento após a inicialização.</span><span class="sxs-lookup"><span data-stu-id="fae55-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fae55-219">O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae55-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="fae55-220">Um erro será acionado se for chamado novamente.</span><span class="sxs-lookup"><span data-stu-id="fae55-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="fae55-221">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae55-221">The following is an example.</span></span> <span data-ttu-id="fae55-222">Observe que a cadeia de caracteres JSON deve ser convertida em um objeto JavaScript com o `JSON.parse` método antes que ele possa ser passado para uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fae55-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ' ... '; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="fae55-223">Especifique os contextos quando a guia estará visível com requestUpdate</span><span class="sxs-lookup"><span data-stu-id="fae55-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="fae55-224">Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fae55-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="fae55-225">Considere um cenário em que a guia deve estar visível quando, e somente quando, um gráfico (na planilha padrão de uma pasta de trabalho do Excel) estiver ativado.</span><span class="sxs-lookup"><span data-stu-id="fae55-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="fae55-226">Comece atribuindo manipuladores.</span><span class="sxs-lookup"><span data-stu-id="fae55-226">Begin by assigning handlers.</span></span> <span data-ttu-id="fae55-227">Isso geralmente é feito no `Office.onReady` método como no exemplo a seguir, que atribui manipuladores (criados em uma etapa posterior) aos `onActivated` `onDeactivated` eventos e de todos os gráficos da planilha.</span><span class="sxs-lookup"><span data-stu-id="fae55-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="fae55-228">Em seguida, defina os manipuladores.</span><span class="sxs-lookup"><span data-stu-id="fae55-228">Next, define the handlers.</span></span> <span data-ttu-id="fae55-229">Veja a seguir um exemplo simples de um `showDataTab` , mas consulte [tratamento de erros](#error-handling) posteriormente neste artigo para obter uma versão mais robusta da função.</span><span class="sxs-lookup"><span data-stu-id="fae55-229">The following is a simple example of a `showDataTab`, but see [Error Handling](#error-handling) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="fae55-230">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="fae55-230">About this code, note:</span></span>

- <span data-ttu-id="fae55-231">O Office controla quando atualiza o estado da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="fae55-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="fae55-232">O método  [Office. Ribbon. requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) enfileira uma solicitação para atualizar.</span><span class="sxs-lookup"><span data-stu-id="fae55-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="fae55-233">O método resolverá o `Promise` objeto assim que ele enfileirar a solicitação, não quando a faixa de opções realmente for atualizada.</span><span class="sxs-lookup"><span data-stu-id="fae55-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="fae55-234">O parâmetro para o `requestUpdate` método é um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a Tabulação por sua ID *exatamente conforme especificado no JSON* e (2) especifica a visibilidade da guia.</span><span class="sxs-lookup"><span data-stu-id="fae55-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="fae55-235">Se você tiver mais de uma guia contextual personalizada que deve estar visível no mesmo contexto, basta adicionar outros objetos Tab à `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="fae55-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="fae55-236">O manipulador para ocultar a guia é quase idêntico, exceto pelo fato de que ela define a `visible` propriedade de volta para `false` .</span><span class="sxs-lookup"><span data-stu-id="fae55-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="fae55-237">A biblioteca JavaScript do Office também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto.</span><span class="sxs-lookup"><span data-stu-id="fae55-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="fae55-238">A seguir está a `showDataTab` função no TypeScript e utiliza esses tipos.</span><span class="sxs-lookup"><span data-stu-id="fae55-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="fae55-239">Alternar a visibilidade da guia e o status habilitado de um botão ao mesmo tempo</span><span class="sxs-lookup"><span data-stu-id="fae55-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="fae55-240">O `requestUpdate` método também é usado para alternar o status habilitado ou desabilitado de um botão personalizado em uma guia contextual personalizada ou em uma guia principal personalizada. Para obter detalhes sobre isso, consulte [habilitar e desabilitar comandos de suplemento](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="fae55-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="fae55-241">Pode haver cenários nos quais você deseja alterar a visibilidade de uma guia e o status habilitado de um botão ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="fae55-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="fae55-242">Você pode fazer isso com uma única chamada de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="fae55-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="fae55-243">A seguir está um exemplo no qual um botão em uma guia principal está habilitado ao mesmo tempo em que uma guia contextual é torna visível.</span><span class="sxs-lookup"><span data-stu-id="fae55-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="fae55-244">No exemplo a seguir, o botão habilitado está na mesma guia contextual que está sendo visível.</span><span class="sxs-lookup"><span data-stu-id="fae55-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="error-handling"></a><span data-ttu-id="fae55-245">Tratamento de erros</span><span class="sxs-lookup"><span data-stu-id="fae55-245">Error handling</span></span>

<span data-ttu-id="fae55-246">Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="fae55-246">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="fae55-247">Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="fae55-247">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="fae55-248">Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="fae55-248">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="fae55-249">Veja um exemplo de como lidar com esse erro a seguir.</span><span class="sxs-lookup"><span data-stu-id="fae55-249">The following is an example of how to handle this error.</span></span> <span data-ttu-id="fae55-250">Nesse caso, o método `reportError` exibe o erro para o usuário.</span><span class="sxs-lookup"><span data-stu-id="fae55-250">In this case, the `reportError` method displays the error to the user.</span></span>

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
