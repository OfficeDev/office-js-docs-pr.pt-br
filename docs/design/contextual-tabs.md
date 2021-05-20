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
# <a name="create-custom-contextual-tabs-in-office-add-ins"></a><span data-ttu-id="ef3de-103">Crie guias contextuais personalizadas em Office Add-ins</span><span class="sxs-lookup"><span data-stu-id="ef3de-103">Create custom contextual tabs in Office Add-ins</span></span>

<span data-ttu-id="ef3de-104">Uma guia contextual é um controle de guia oculto na fita Office que é exibida na linha de guia quando um evento especificado ocorre no documento Office.</span><span class="sxs-lookup"><span data-stu-id="ef3de-104">A contextual tab is a hidden tab control in the Office ribbon that is displayed in the tab row when a specified event occurs in the Office document.</span></span> <span data-ttu-id="ef3de-105">Por exemplo, a guia **Design de tabela** que aparece na fita Excel quando uma tabela é selecionada.</span><span class="sxs-lookup"><span data-stu-id="ef3de-105">For example, the **Table Design** tab that appears on the Excel ribbon when a table is selected.</span></span> <span data-ttu-id="ef3de-106">Você pode incluir guias contextuais personalizadas em seu Office Add-in e especificar quando elas estão visíveis ou ocultas, criando manipuladores de eventos que alteram a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="ef3de-106">You can include custom contextual tabs in your Office Add-in and specify when they are visible or hidden, by creating event handlers that change the visibility.</span></span> <span data-ttu-id="ef3de-107">(No entanto, as guias contextuais personalizadas não respondem às alterações de foco.)</span><span class="sxs-lookup"><span data-stu-id="ef3de-107">(However, custom contextual tabs do not respond to focus changes.)</span></span>

> [!NOTE]
> <span data-ttu-id="ef3de-108">Este artigo pressupõe que você esteja familiarizado com a seguinte documentação.</span><span class="sxs-lookup"><span data-stu-id="ef3de-108">This article assumes that you are familiar with the following documentation.</span></span> <span data-ttu-id="ef3de-109">Revise-o se você não trabalhou recentemente com os Comandos de Suplemento (itens de menu personalizados e botões da faixa de opções).</span><span class="sxs-lookup"><span data-stu-id="ef3de-109">Please review it if you haven't worked with Add-in Commands (custom menu items and ribbon buttons) recently.</span></span>
>
> - [<span data-ttu-id="ef3de-110">Conceitos básicos dos Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="ef3de-110">Basic concepts for Add-in Commands</span></span>](add-in-commands.md)

> [!IMPORTANT]
> <span data-ttu-id="ef3de-111">Atualmente, as guias contextuais personalizadas são suportadas apenas em Excel e somente nessas plataformas e compilações:</span><span class="sxs-lookup"><span data-stu-id="ef3de-111">Custom contextual tabs are currently only supported on Excel and only on these platforms and builds:</span></span>
>
> - <span data-ttu-id="ef3de-112">Excel em Windows (somente Microsoft 365 assinatura): Versão 2102 (Build 13801.20294) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="ef3de-112">Excel on Windows (Microsoft 365 subscription only): Version 2102 (Build 13801.20294) or later.</span></span>
> - <span data-ttu-id="ef3de-113">Excel Online</span><span class="sxs-lookup"><span data-stu-id="ef3de-113">Excel on the web</span></span>

> [!NOTE]
> <span data-ttu-id="ef3de-114">As guias contextuais personalizadas funcionam apenas em plataformas que suportam os seguintes conjuntos de requisitos.</span><span class="sxs-lookup"><span data-stu-id="ef3de-114">Custom contextual tabs work only on platforms that support the following requirement sets.</span></span> <span data-ttu-id="ef3de-115">Para obter mais informações sobre os conjuntos de requisitos e como trabalhar com eles, consulte [Especificar Office aplicativos e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-115">For more about requirement sets and how to work with them, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md).</span></span>
>
> - [<span data-ttu-id="ef3de-116">RibbonApi 1.2</span><span class="sxs-lookup"><span data-stu-id="ef3de-116">RibbonApi 1.2</span></span>](../reference/requirement-sets/ribbon-api-requirement-sets.md)
> - [<span data-ttu-id="ef3de-117">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="ef3de-117">SharedRuntime 1.1</span></span>](../reference/requirement-sets/shared-runtime-requirement-sets.md)
>
> <span data-ttu-id="ef3de-118">Você pode usar as verificações de tempo de execução em seu código para testar se a combinação de host e plataforma do usuário suporta esses conjuntos de requisitos conforme descrito em [Especificar Office aplicativos e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span><span class="sxs-lookup"><span data-stu-id="ef3de-118">You can use the runtime checks in your code to test whether the user's host and platform combination supports these requirement sets as described in [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code).</span></span> <span data-ttu-id="ef3de-119">(A técnica de especificar os conjuntos de exigências no manifesto, que também está descrito nesse artigo, não funciona atualmente para RibbonApi 1.2.) Alternativamente, você pode [implementar uma experiência de interface do usuário alternativa quando as guias contextuais personalizadas não forem suportadas](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span><span class="sxs-lookup"><span data-stu-id="ef3de-119">(The technique of specifying the requirement sets in the manifest, which is also described in that article, does not currently work for RibbonApi 1.2.) Alternatively, you can [implement an alternate UI experience when custom contextual tabs are not supported](#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).</span></span>

## <a name="behavior-of-custom-contextual-tabs"></a><span data-ttu-id="ef3de-120">Comportamento de guias contextuais personalizadas</span><span class="sxs-lookup"><span data-stu-id="ef3de-120">Behavior of custom contextual tabs</span></span>

<span data-ttu-id="ef3de-121">A experiência do usuário para guias contextuais personalizadas segue o padrão de guias Office contextuais incorporadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-121">The user experience for custom contextual tabs follows the pattern of built-in Office contextual tabs.</span></span> <span data-ttu-id="ef3de-122">A seguir, os princípios básicos para as guias contextuais personalizadas de colocação:</span><span class="sxs-lookup"><span data-stu-id="ef3de-122">The following are the basic principles for the placement custom contextual tabs:</span></span>

- <span data-ttu-id="ef3de-123">Quando uma guia contextual personalizada é visível, ela aparece na extremidade direita da fita.</span><span class="sxs-lookup"><span data-stu-id="ef3de-123">When a custom contextual tab is visible, it appears on the right end of the ribbon.</span></span>
- <span data-ttu-id="ef3de-124">Se uma ou mais guias contextuais incorporadas e uma ou mais guias contextuais personalizadas de complementos forem visíveis ao mesmo tempo, as guias contextuais personalizadas estão sempre à direita de todas as guias contextuais incorporadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-124">If one or more built-in contextual tabs and one or more custom contextual tabs from add-ins are visible at the same time, the custom contextual tabs are always to the right of all of the built-in contextual tabs.</span></span>
- <span data-ttu-id="ef3de-125">Se o seu complemento tiver mais de uma guia contextual e houver contextos em que mais de um seja visível, eles aparecem na ordem em que são definidos em seu complemento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-125">If your add-in has more than one contextual tab and there are contexts in which more than one is visible, they appear in the order in which they are defined in your add-in.</span></span> <span data-ttu-id="ef3de-126">(A direção é a mesma direção que a língua Office; ou seja, é da esquerda para a direita em línguas da esquerda para a direita, mas da direita para a esquerda em línguas da direita para a esquerda.) Consulte [Definir os grupos e controles que aparecem na guia](#define-the-groups-and-controls-that-appear-on-the-tab) para obter detalhes sobre como você os define.</span><span class="sxs-lookup"><span data-stu-id="ef3de-126">(The direction is the same direction as the Office language; that is, is left-to-right in left-to-right languages, but right-to-left in right-to-left languages.) See [Define the groups and controls that appear on the tab](#define-the-groups-and-controls-that-appear-on-the-tab) for details about how you define them.</span></span>
- <span data-ttu-id="ef3de-127">Se mais de um complemento tiver uma guia contextual visível em um contexto específico, então eles aparecem na ordem em que os complementos foram lançados.</span><span class="sxs-lookup"><span data-stu-id="ef3de-127">If more than one add-in has a contextual tab that is visible in a specific context, then they appear in the order in which the add-ins were launched.</span></span>
- <span data-ttu-id="ef3de-128">As guias *contextuais* personalizadas, ao contrário das guias de núcleo personalizadas, não são adicionadas permanentemente à fita do aplicativo Office.</span><span class="sxs-lookup"><span data-stu-id="ef3de-128">Custom *contextual* tabs, unlike custom core tabs, are not added permanently to the Office application's ribbon.</span></span> <span data-ttu-id="ef3de-129">Eles estão presentes apenas em Office documentos sobre os quais seu complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-129">They are present only in Office documents on which your add-in is running.</span></span>

## <a name="major-steps-for-including-a-contextual-tab-in-an-add-in"></a><span data-ttu-id="ef3de-130">Principais etapas para incluir uma guia contextual em um complemento</span><span class="sxs-lookup"><span data-stu-id="ef3de-130">Major steps for including a contextual tab in an add-in</span></span>

<span data-ttu-id="ef3de-131">A seguir, os principais passos para incluir uma guia contextual personalizada em um complemento:</span><span class="sxs-lookup"><span data-stu-id="ef3de-131">The following are the major steps for including a custom contextual tab in an add-in:</span></span>

1. <span data-ttu-id="ef3de-132">Configure o complemento para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-132">Configure the add-in to use a shared runtime.</span></span>
1. <span data-ttu-id="ef3de-133">Defina a guia e os grupos e controles que aparecem nela.</span><span class="sxs-lookup"><span data-stu-id="ef3de-133">Define the tab and the groups and controls that appear on it.</span></span>
1. <span data-ttu-id="ef3de-134">Cadastre-se na aba contextual com Office.</span><span class="sxs-lookup"><span data-stu-id="ef3de-134">Register the contextual tab with Office.</span></span>
1. <span data-ttu-id="ef3de-135">Especifique as circunstâncias em que a guia será visível.</span><span class="sxs-lookup"><span data-stu-id="ef3de-135">Specify the circumstances when the tab will be visible.</span></span>

## <a name="configure-the-add-in-to-use-a-shared-runtime"></a><span data-ttu-id="ef3de-136">Configure o complemento para usar um tempo de execução compartilhado</span><span class="sxs-lookup"><span data-stu-id="ef3de-136">Configure the add-in to use a shared runtime</span></span>

<span data-ttu-id="ef3de-137">Adicionar guias contextuais personalizadas requer que seu complemento use o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-137">Adding custom contextual tabs requires your add-in to use the shared runtime.</span></span> <span data-ttu-id="ef3de-138">Para obter mais informações, consulte [Configurar um complemento para usar um tempo de execução compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-138">For more information, see [Configure an add-in to use a shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="define-the-groups-and-controls-that-appear-on-the-tab"></a><span data-ttu-id="ef3de-139">Defina os grupos e controles que aparecem na guia</span><span class="sxs-lookup"><span data-stu-id="ef3de-139">Define the groups and controls that appear on the tab</span></span>

<span data-ttu-id="ef3de-140">Ao contrário das guias de núcleo personalizadas, que são definidas com XML no manifesto, as guias contextuais personalizadas são definidas no tempo de execução com uma bolha JSON.</span><span class="sxs-lookup"><span data-stu-id="ef3de-140">Unlike custom core tabs, which are defined with XML in the manifest, custom contextual tabs are defined at runtime with a JSON blob.</span></span> <span data-ttu-id="ef3de-141">Seu código analisa a bolha em um objeto JavaScript e, em seguida, passa o objeto para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-)</span><span class="sxs-lookup"><span data-stu-id="ef3de-141">Your code parses the blob into a JavaScript object, and then passes the object to the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls-tabDefinition-) method.</span></span> <span data-ttu-id="ef3de-142">As guias contextuais personalizadas só estão presentes em documentos nos quais seu complemento está sendo executado no momento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-142">Custom contextual tabs are only present in documents on which your add-in is currently running.</span></span> <span data-ttu-id="ef3de-143">Isso é diferente das guias de núcleo personalizadas que são adicionadas à fita de aplicação Office quando o complemento é instalado e permanecem presentes quando outro documento é aberto.</span><span class="sxs-lookup"><span data-stu-id="ef3de-143">This is different from custom core tabs which are added to the Office application ribbon when the add-in is installed and remain present when another document is opened.</span></span> <span data-ttu-id="ef3de-144">Além disso, o `requestCreateControls` método pode ser executado apenas uma vez em uma sessão do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-144">Also, the `requestCreateControls` method can be run only once in a session of your add-in.</span></span> <span data-ttu-id="ef3de-145">Se for chamado novamente, um erro é jogado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-145">If it is called again, an error is thrown.</span></span>

> [!NOTE]
> <span data-ttu-id="ef3de-146">A estrutura das propriedades e subpropriedades do blob JSON (e os nomes-chave) é aproximadamente paralela à estrutura do elemento [CustomTab](../reference/manifest/customtab.md) e seus elementos descendentes no manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="ef3de-146">The structure of the JSON blob's properties and subproperties (and the key names) is roughly parallel to the structure of the [CustomTab](../reference/manifest/customtab.md) element and its descendant elements in the manifest XML.</span></span>

<span data-ttu-id="ef3de-147">Construiremos um exemplo de uma aba contextual JSON blob passo a passo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-147">We'll construct an example of a contextual tabs JSON blob step-by-step.</span></span> <span data-ttu-id="ef3de-148">O esquema completo para a aba contextual JSON está em [dynamic-ribbon.schema.jsem.](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json)</span><span class="sxs-lookup"><span data-stu-id="ef3de-148">The full schema for the contextual tab JSON is at [dynamic-ribbon.schema.json](https://developer.microsoft.com/json-schemas/office-js/dynamic-ribbon.schema.json).</span></span> <span data-ttu-id="ef3de-149">Se você estiver trabalhando em Visual Studio Code, você pode usar este arquivo para obter IntelliSense e validar seu JSON.</span><span class="sxs-lookup"><span data-stu-id="ef3de-149">If you are working in Visual Studio Code, you can use this file to get IntelliSense and to validate your JSON.</span></span> <span data-ttu-id="ef3de-150">Para obter mais informações, consulte [Editando JSON com Visual Studio Code - esquemas e configurações JSON](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span><span class="sxs-lookup"><span data-stu-id="ef3de-150">For more information, see [Editing JSON with Visual Studio Code - JSON schemas and settings](https://code.visualstudio.com/docs/languages/json#_json-schemas-and-settings).</span></span>


1. <span data-ttu-id="ef3de-151">Comece criando uma sequência JSON com duas propriedades de matriz `actions` nomeadas e `tabs` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-151">Begin by creating a JSON string with two array properties named `actions` and `tabs`.</span></span> <span data-ttu-id="ef3de-152">O `actions` array é uma especificação de todas as funções que podem ser executadas por controles na guia contextual. A `tabs` matriz define uma ou mais guias contextuais, até um máximo de *20*.</span><span class="sxs-lookup"><span data-stu-id="ef3de-152">The `actions` array is a specification of all the functions that can be executed by controls on the contextual tab. The `tabs` array defines one or more contextual tabs, *up to a maximum of 20*.</span></span>

    ```json
    '{
      "actions": [

      ],
      "tabs": [

      ]
    }'
    ```

1. <span data-ttu-id="ef3de-153">Este simples exemplo de uma guia contextual terá apenas um único botão e, portanto, apenas uma única ação.</span><span class="sxs-lookup"><span data-stu-id="ef3de-153">This simple example of a contextual tab will have only a single button and, thus, only a single action.</span></span> <span data-ttu-id="ef3de-154">Adicione o seguinte como o único membro da `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-154">Add the following as the only member of the `actions` array.</span></span> <span data-ttu-id="ef3de-155">Sobre esta marcação, nota:</span><span class="sxs-lookup"><span data-stu-id="ef3de-155">About this markup, note:</span></span>

    - <span data-ttu-id="ef3de-156">As `id` `type` propriedades são obrigatórias.</span><span class="sxs-lookup"><span data-stu-id="ef3de-156">The `id` and `type` properties are mandatory.</span></span>
    - <span data-ttu-id="ef3de-157">O valor `type` pode ser "ExecuteFunction" ou "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="ef3de-157">The value of `type` can be either "ExecuteFunction" or "ShowTaskpane".</span></span>
    - <span data-ttu-id="ef3de-158">O `functionName` imóvel só é usado quando o valor é `type` `ExecuteFunction` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-158">The `functionName` property is only used when the value of `type` is `ExecuteFunction`.</span></span> <span data-ttu-id="ef3de-159">É o nome de uma função definida no FunctionFile.</span><span class="sxs-lookup"><span data-stu-id="ef3de-159">It is the name of a function defined in the FunctionFile.</span></span> <span data-ttu-id="ef3de-160">Para obter mais informações sobre o FunctionFile, consulte [conceitos básicos para comandos adicionais](add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-160">For more information about the FunctionFile, see [Basic concepts for Add-in Commands](add-in-commands.md).</span></span>
    - <span data-ttu-id="ef3de-161">Em um passo posterior, você mapeará esta ação para um botão na guia contextual.</span><span class="sxs-lookup"><span data-stu-id="ef3de-161">In a later step, you will map this action to a button on the contextual tab.</span></span>

    ```json
    {
      "id": "executeWriteData",
      "type": "ExecuteFunction",
      "functionName": "writeData"
    }
   ```

1. <span data-ttu-id="ef3de-162">Adicione o seguinte como o único membro da `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-162">Add the following as the only member of the `tabs` array.</span></span> <span data-ttu-id="ef3de-163">Sobre esta marcação, nota:</span><span class="sxs-lookup"><span data-stu-id="ef3de-163">About this markup, note:</span></span>

    - <span data-ttu-id="ef3de-164">A propriedade `id` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="ef3de-164">The `id` property is required.</span></span> <span data-ttu-id="ef3de-165">Use um ID breve e descritivo que seja único entre todas as guias contextuais em seu complemento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-165">Use a brief, descriptive ID that is unique among all contextual tabs in your add-in.</span></span>
    - <span data-ttu-id="ef3de-166">A propriedade `label` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="ef3de-166">The `label` property is required.</span></span> <span data-ttu-id="ef3de-167">É uma sequência fácil de usar para servir como o rótulo da guia contextual.</span><span class="sxs-lookup"><span data-stu-id="ef3de-167">It is a user-friendly string to serve as the label of the contextual tab.</span></span>
    - <span data-ttu-id="ef3de-168">A propriedade `groups` é obrigatória.</span><span class="sxs-lookup"><span data-stu-id="ef3de-168">The `groups` property is required.</span></span> <span data-ttu-id="ef3de-169">Ele define os grupos de controles que aparecerão na guia. Deve ter pelo menos um membro *e não mais de 20.*</span><span class="sxs-lookup"><span data-stu-id="ef3de-169">It defines the groups of controls that will appear on the tab. It must have at least one member *and no more than 20*.</span></span> <span data-ttu-id="ef3de-170">(Há também limites no número de controles que você pode ter em uma guia contextual personalizada e isso também restringirá quantos grupos você tem.</span><span class="sxs-lookup"><span data-stu-id="ef3de-170">(There are also limits on the number of controls that you can have on a custom contextual tab and that will also constrain how many groups that you have.</span></span> <span data-ttu-id="ef3de-171">Veja o próximo passo para obter mais informações.)</span><span class="sxs-lookup"><span data-stu-id="ef3de-171">See the next step for more information.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="ef3de-172">O objeto da guia também pode ter uma propriedade opcional `visible` que especifica se a guia é visível imediatamente quando o complemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-172">The tab object can also have an optional `visible` property that specifies whether the tab is visible immediately when the add-in starts up.</span></span> <span data-ttu-id="ef3de-173">Uma vez que as guias contextuais são normalmente ocultas até que um evento de usuário acione sua visibilidade (como o usuário selecionando uma entidade de algum tipo no documento), o `visible` imóvel padrão para quando não está `false` presente.</span><span class="sxs-lookup"><span data-stu-id="ef3de-173">Since contextual tabs are normally hidden until a user event triggers their visibility (such as the user selecting an entity of some type in the document), the `visible` property defaults to `false` when not present.</span></span> <span data-ttu-id="ef3de-174">Em uma seção posterior, mostramos como definir a propriedade `true` em resposta a um evento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-174">In a later section, we show how to set the property to `true` in response to an event.</span></span>

    ```json
    {
      "id": "CtxTab1",
      "label": "Contoso Data",
      "groups": [

      ]
    }
    ```

1. <span data-ttu-id="ef3de-175">No simples exemplo em curso, a guia contextual tem apenas um único grupo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-175">In the simple ongoing example, the contextual tab has only a single group.</span></span> <span data-ttu-id="ef3de-176">Adicione o seguinte como o único membro da `groups` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-176">Add the following as the only member of the `groups` array.</span></span> <span data-ttu-id="ef3de-177">Sobre esta marcação, nota:</span><span class="sxs-lookup"><span data-stu-id="ef3de-177">About this markup, note:</span></span>

    - <span data-ttu-id="ef3de-178">Todas as propriedades são necessárias.</span><span class="sxs-lookup"><span data-stu-id="ef3de-178">All the properties are required.</span></span>
    - <span data-ttu-id="ef3de-179">A `id` propriedade deve ser única entre todos os grupos da guia. Use um Y breve e descritivo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-179">The `id` property must be unique among all the groups in the tab. Use a brief, descriptive ID.</span></span>
    - <span data-ttu-id="ef3de-180">A `label` é uma string fácil de usar para servir como o rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-180">The `label` is a user-friendly string to serve as the label of the group.</span></span>
    - <span data-ttu-id="ef3de-181">O `icon` valor da propriedade é uma matriz de objetos que especificam os ícones que o grupo terá na fita, dependendo do tamanho da fita e da janela de aplicação Office.</span><span class="sxs-lookup"><span data-stu-id="ef3de-181">The `icon` property's value is an array of objects that specify the icons that the group will have on the ribbon depending on the size of the ribbon and the Office application window.</span></span>
    - <span data-ttu-id="ef3de-182">O `controls` valor da propriedade é uma matriz de objetos que especificam os botões e menus do grupo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-182">The `controls` property's value is an array of objects that specify the buttons and menus in the group.</span></span> <span data-ttu-id="ef3de-183">Deve haver pelo menos um.</span><span class="sxs-lookup"><span data-stu-id="ef3de-183">There must be at least one.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ef3de-184">*O número total de controles na guia geral não pode ser superior a 20.*</span><span class="sxs-lookup"><span data-stu-id="ef3de-184">*The total number of controls on the whole tab can be no more than 20.*</span></span> <span data-ttu-id="ef3de-185">Por exemplo, você poderia ter 3 grupos com 6 controles cada, e um quarto grupo com 2 controles, mas você não pode ter 4 grupos com 6 controles cada.</span><span class="sxs-lookup"><span data-stu-id="ef3de-185">For example, you could have 3 groups with 6 controls each, and a fourth group with 2 controls, but you cannot have 4 groups with 6 controls each.</span></span>  

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

1. <span data-ttu-id="ef3de-186">Cada grupo deve ter um ícone de pelo menos dois tamanhos, 32x32 px e 80x80 px.</span><span class="sxs-lookup"><span data-stu-id="ef3de-186">Every group must have an icon of at least two sizes, 32x32 px and 80x80 px.</span></span> <span data-ttu-id="ef3de-187">Opcionalmente, você também pode ter ícones dos tamanhos 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px e 64x64 px.</span><span class="sxs-lookup"><span data-stu-id="ef3de-187">Optionally, you can also have icons of sizes 16x16 px, 20x20 px, 24x24 px, 40x40 px, 48x48 px, and 64x64 px.</span></span> <span data-ttu-id="ef3de-188">Office decide qual ícone usar com base no tamanho da fita e Office janela de aplicativo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-188">Office decides which icon to use based on the size of the ribbon and Office application window.</span></span> <span data-ttu-id="ef3de-189">Adicione os seguintes objetos à matriz de ícones.</span><span class="sxs-lookup"><span data-stu-id="ef3de-189">Add the following objects to the icon array.</span></span> <span data-ttu-id="ef3de-190">(Se os tamanhos da janela e da fita forem grandes o suficiente para que pelo menos um dos *controles* do grupo apareça, então nenhum ícone de grupo aparece.</span><span class="sxs-lookup"><span data-stu-id="ef3de-190">(If the window and ribbon sizes are large enough for at least one of the *controls* on the group to appear, then no group icon at all appears.</span></span> <span data-ttu-id="ef3de-191">Por exemplo, assista ao grupo **Styles** na fita do Word enquanto você encolhe e expande a janela do Word.) Sobre esta marcação, nota:</span><span class="sxs-lookup"><span data-stu-id="ef3de-191">For an example, watch the **Styles** group on the Word ribbon as you shrink and expand the Word window.) About this markup, note:</span></span>

    - <span data-ttu-id="ef3de-192">Ambas as propriedades são necessárias.</span><span class="sxs-lookup"><span data-stu-id="ef3de-192">Both the properties are required.</span></span>
    - <span data-ttu-id="ef3de-193">A `size` unidade de propriedade da medida é pixels.</span><span class="sxs-lookup"><span data-stu-id="ef3de-193">The `size` property unit of measure is pixels.</span></span> <span data-ttu-id="ef3de-194">Os ícones são sempre quadrados, então o número é tanto a altura quanto a largura.</span><span class="sxs-lookup"><span data-stu-id="ef3de-194">Icons are always square, so the number is both the height and the width.</span></span>
    - <span data-ttu-id="ef3de-195">A `sourceLocation` propriedade especifica a URL completa para o ícone.</span><span class="sxs-lookup"><span data-stu-id="ef3de-195">The `sourceLocation` property specifies the full URL to the icon.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="ef3de-196">Assim como você normalmente deve alterar os URLs no manifesto do complemento quando você passar do desenvolvimento para a produção (como mudar o domínio de localhost para contoso.com), você também deve alterar as URLs em suas guias contextuais JSON.</span><span class="sxs-lookup"><span data-stu-id="ef3de-196">Just as you typically must change the URLs in the add-in's manifest when you move from development to production (such as changing the domain from localhost to contoso.com), you must also change the URLs in your contextual tabs JSON.</span></span>

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

1. <span data-ttu-id="ef3de-197">Em nosso simples exemplo contínuo, o grupo tem apenas um único botão.</span><span class="sxs-lookup"><span data-stu-id="ef3de-197">In our simple ongoing example, the group has only a single button.</span></span> <span data-ttu-id="ef3de-198">Adicione o seguinte objeto como o único membro da `controls` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-198">Add the following object as the only member of the `controls` array.</span></span> <span data-ttu-id="ef3de-199">Sobre esta marcação, nota:</span><span class="sxs-lookup"><span data-stu-id="ef3de-199">About this markup, note:</span></span>

    - <span data-ttu-id="ef3de-200">Todas as propriedades, `enabled` exceto, são necessárias.</span><span class="sxs-lookup"><span data-stu-id="ef3de-200">All the properties, except `enabled`, are required.</span></span>
    - <span data-ttu-id="ef3de-201">`type` especifica o tipo de controle.</span><span class="sxs-lookup"><span data-stu-id="ef3de-201">`type` specifies the type of control.</span></span> <span data-ttu-id="ef3de-202">Os valores podem ser "Button", "Menu" ou "MobileButton".</span><span class="sxs-lookup"><span data-stu-id="ef3de-202">The values can be "Button", "Menu", or "MobileButton".</span></span>
    - <span data-ttu-id="ef3de-203">`id` pode ter até 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ef3de-203">`id` can be up to 125 characters.</span></span> 
    - <span data-ttu-id="ef3de-204">`actionId` deve ser o ID de uma ação definida na `actions` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-204">`actionId` must be the ID of an action defined in the `actions` array.</span></span> <span data-ttu-id="ef3de-205">(Veja o passo 1 desta seção.)</span><span class="sxs-lookup"><span data-stu-id="ef3de-205">(See step 1 of this section.)</span></span>
    - <span data-ttu-id="ef3de-206">`label` é uma string fácil de usar para servir como a etiqueta do botão.</span><span class="sxs-lookup"><span data-stu-id="ef3de-206">`label` is a user-friendly string to serve as the label of the button.</span></span>
    - <span data-ttu-id="ef3de-207">`superTip` representa uma rica forma de ponta de ferramenta.</span><span class="sxs-lookup"><span data-stu-id="ef3de-207">`superTip` represents a rich form of tool tip.</span></span> <span data-ttu-id="ef3de-208">Tanto as propriedades quanto as `title` `description` propriedades são necessárias.</span><span class="sxs-lookup"><span data-stu-id="ef3de-208">Both the `title` and `description` properties are required.</span></span>
    - <span data-ttu-id="ef3de-209">`icon` especifica os ícones para o botão.</span><span class="sxs-lookup"><span data-stu-id="ef3de-209">`icon` specifies the icons for the button.</span></span> <span data-ttu-id="ef3de-210">As observações anteriores sobre o ícone de grupo também se aplicam aqui.</span><span class="sxs-lookup"><span data-stu-id="ef3de-210">The previous remarks about the group icon apply here too.</span></span>
    - <span data-ttu-id="ef3de-211">`enabled` (opcional) especifica se o botão está ativado quando a guia contextual é ativada.</span><span class="sxs-lookup"><span data-stu-id="ef3de-211">`enabled` (optional) specifies whether the button is enabled when the contextual tab appears starts up.</span></span> <span data-ttu-id="ef3de-212">O padrão se não estiver presente é `true` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-212">The default if not present is `true`.</span></span> 

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
 
<span data-ttu-id="ef3de-213">O seguinte é o exemplo completo da bolha JSON:</span><span class="sxs-lookup"><span data-stu-id="ef3de-213">The following is the complete example of the JSON blob:</span></span>

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

## <a name="register-the-contextual-tab-with-office-with-requestcreatecontrols"></a><span data-ttu-id="ef3de-214">Cadastre-se na guia contextual com Office solicitaçãoCreateControls</span><span class="sxs-lookup"><span data-stu-id="ef3de-214">Register the contextual tab with Office with requestCreateControls</span></span>

<span data-ttu-id="ef3de-215">A guia contextual é registrada com Office ligando para o método [Office.ribbon.requestCreateControls.](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_)</span><span class="sxs-lookup"><span data-stu-id="ef3de-215">The contextual tab is registered with Office by calling the [Office.ribbon.requestCreateControls](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestCreateControls_tabDefinition_) method.</span></span> <span data-ttu-id="ef3de-216">Isso é normalmente feito na função que é atribuída `Office.initialize` ou com o `Office.onReady` método.</span><span class="sxs-lookup"><span data-stu-id="ef3de-216">This is typically done in either the function that is assigned to `Office.initialize` or with the `Office.onReady` method.</span></span> <span data-ttu-id="ef3de-217">Para obter mais informações sobre esses métodos e inicializar o complemento, consulte [Initialize seu Office Add-in](../develop/initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-217">For more about these methods and initializing the add-in, see [Initialize your Office Add-in](../develop/initialize-add-in.md).</span></span> <span data-ttu-id="ef3de-218">Você pode, no entanto, chamar o método a qualquer momento após a inicialização.</span><span class="sxs-lookup"><span data-stu-id="ef3de-218">You can, however, call the method anytime after initialization.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef3de-219">O `requestCreateControls` método pode ser chamado apenas uma vez em uma determinada sessão de um complemento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-219">The `requestCreateControls` method can be called only once in a given session of an add-in.</span></span> <span data-ttu-id="ef3de-220">Um erro é jogado se for chamado novamente.</span><span class="sxs-lookup"><span data-stu-id="ef3de-220">An error is thrown if it is called again.</span></span>

<span data-ttu-id="ef3de-221">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ef3de-221">The following is an example.</span></span> <span data-ttu-id="ef3de-222">Observe que a sequência JSON deve ser convertida em um objeto JavaScript com o `JSON.parse` método antes que possa ser passado para uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ef3de-222">Note that the JSON string must be converted to a JavaScript object with the `JSON.parse` method before it can be passed to a JavaScript function.</span></span>

```javascript
Office.onReady(async () => {
    const contextualTabJSON = ` ... `; // Assign the JSON string such as the one at the end of the preceding section.
    const contextualTab = JSON.parse(contextualTabJSON);
    await Office.ribbon.requestCreateControls(contextualTab);
});
```

## <a name="specify-the-contexts-when-the-tab-will-be-visible-with-requestupdate"></a><span data-ttu-id="ef3de-223">Especifique os contextos quando a guia estará visível com a solicitaçãoUpdate</span><span class="sxs-lookup"><span data-stu-id="ef3de-223">Specify the contexts when the tab will be visible with requestUpdate</span></span>

<span data-ttu-id="ef3de-224">Normalmente, uma guia contextual personalizada deve aparecer quando um evento iniciado pelo usuário altera o contexto de complementação.</span><span class="sxs-lookup"><span data-stu-id="ef3de-224">Typically, a custom contextual tab should appear when a user-initiated event changes the add-in context.</span></span> <span data-ttu-id="ef3de-225">Considere um cenário em que a guia deve ser visível quando, e somente quando, um gráfico (na planilha padrão de uma Excel pasta de trabalho) é ativado.</span><span class="sxs-lookup"><span data-stu-id="ef3de-225">Consider a scenario in which the tab should be visible when, and only when, a chart (on the default worksheet of an Excel workbook) is activated.</span></span>

<span data-ttu-id="ef3de-226">Comece designando manipuladores.</span><span class="sxs-lookup"><span data-stu-id="ef3de-226">Begin by assigning handlers.</span></span> <span data-ttu-id="ef3de-227">Isso é comumente feito no `Office.onReady` método como no exemplo a seguir que atribui manipuladores (criados em uma etapa posterior) aos eventos `onActivated` de todos os `onDeactivated` gráficos na planilha.</span><span class="sxs-lookup"><span data-stu-id="ef3de-227">This is commonly done in the `Office.onReady` method as in the following example which assigns handlers (created in a later step) to the `onActivated` and `onDeactivated` events of all the charts in the worksheet.</span></span>

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

<span data-ttu-id="ef3de-228">Em seguida, defina os manipuladores.</span><span class="sxs-lookup"><span data-stu-id="ef3de-228">Next, define the handlers.</span></span> <span data-ttu-id="ef3de-229">A seguir, um exemplo simples de um `showDataTab` , mas veja [Manipulação do erro HostRestartNeed](#handle-the-hostrestartneeded-error) mais tarde neste artigo para uma versão mais robusta da função.</span><span class="sxs-lookup"><span data-stu-id="ef3de-229">The following is a simple example of a `showDataTab`, but see [Handling the HostRestartNeeded error](#handle-the-hostrestartneeded-error) later in this article for a more robust version of the function.</span></span> <span data-ttu-id="ef3de-230">Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="ef3de-230">About this code, note:</span></span>

- <span data-ttu-id="ef3de-231">O Office controla quando atualiza o estado da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ef3de-231">Office controls when it updates the state of the ribbon.</span></span> <span data-ttu-id="ef3de-232">O método [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) faz filas em uma solicitação para atualizar.</span><span class="sxs-lookup"><span data-stu-id="ef3de-232">The  [Office.ribbon.requestUpdate](/javascript/api/office/office.ribbon?view=common-js&preserve-view=true#requestupdate-input-) method queues a request to update.</span></span> <span data-ttu-id="ef3de-233">O método resolverá o `Promise` objeto assim que ele tiver enfileido a solicitação, não quando a fita realmente se atualizar.</span><span class="sxs-lookup"><span data-stu-id="ef3de-233">The method will resolve the `Promise` object as soon as it has queued the request, not when the ribbon actually updates.</span></span>
- <span data-ttu-id="ef3de-234">O parâmetro para o `requestUpdate` método é um objeto [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) que (1) especifica a guia pelo seu ID *exatamente conforme especificado no JSON* e (2) especifica a visibilidade da guia.</span><span class="sxs-lookup"><span data-stu-id="ef3de-234">The parameter for the `requestUpdate` method is a [RibbonUpdaterData](/javascript/api/office/office.ribbonupdaterdata) object that (1) specifies the tab by its ID *exactly as specified in the JSON* and (2) specifies visibility of the tab.</span></span>
- <span data-ttu-id="ef3de-235">Se você tiver mais de uma guia contextual personalizada que deve ser visível no mesmo contexto, basta adicionar objetos de guia adicionais à `tabs` matriz.</span><span class="sxs-lookup"><span data-stu-id="ef3de-235">If you have more than one custom contextual tab that should be visible in the same context, you simply add additional tab objects to the `tabs` array.</span></span>

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

<span data-ttu-id="ef3de-236">O manipulador para ocultar a guia é quase idêntico, exceto que ele define a `visible` propriedade de volta para `false` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-236">The handler to hide the tab is nearly identical, except that it sets the `visible` property back to `false`.</span></span>

<span data-ttu-id="ef3de-237">A biblioteca Office JavaScript também fornece várias interfaces (tipos) para facilitar a construção do `RibbonUpdateData` objeto.</span><span class="sxs-lookup"><span data-stu-id="ef3de-237">The Office JavaScript library also provides several interfaces (types) to make it easier to construct the`RibbonUpdateData` object.</span></span> <span data-ttu-id="ef3de-238">A seguir, a `showDataTab` função no TypeScript e faz uso desses tipos.</span><span class="sxs-lookup"><span data-stu-id="ef3de-238">The following is the `showDataTab` function in TypeScript and it makes use of these types.</span></span>

```typescript
const showDataTab = async () => {
    const myContextualTab: Office.Tab = {id: "CtxTab1", visible: true};
    const ribbonUpdater: Office.RibbonUpdaterData = { tabs: [ myContextualTab ]};
    await Office.ribbon.requestUpdate(ribbonUpdater);
}
```

### <a name="toggle-tab-visibility-and-the-enabled-status-of-a-button-at-the-same-time"></a><span data-ttu-id="ef3de-239">Alternar a visibilidade da guia e o status habilitado de um botão ao mesmo tempo</span><span class="sxs-lookup"><span data-stu-id="ef3de-239">Toggle tab visibility and the enabled status of a button at the same time</span></span>

<span data-ttu-id="ef3de-240">O `requestUpdate` método também é usado para alternar o status ativado ou desativado de um botão personalizado em uma guia contextual personalizada ou em uma guia central personalizada. Para obter detalhes sobre isso, consulte [Ativar e Desativar comandos adicionais](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-240">The `requestUpdate` method is also used to toggle the enabled or disabled status of a custom button on either a custom contextual tab or a custom core tab. For details about this, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span> <span data-ttu-id="ef3de-241">Pode haver cenários em que você deseja alterar tanto a visibilidade de uma guia quanto o status habilitado de um botão ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-241">There may be scenarios in which you want to change both the visibility of a tab and the enabled status of a button at the same time.</span></span> <span data-ttu-id="ef3de-242">Você pode fazer isso com uma única chamada de `requestUpdate` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-242">You can do this with a single call of `requestUpdate`.</span></span> <span data-ttu-id="ef3de-243">A seguir, um exemplo no qual um botão em uma guia central é ativado ao mesmo tempo que uma guia contextual é visível.</span><span class="sxs-lookup"><span data-stu-id="ef3de-243">The following is an example in which a button on a core tab is enabled at the same time as a contextual tab is made visible.</span></span>

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

<span data-ttu-id="ef3de-244">No exemplo a seguir, o botão habilitado está na mesma aba contextual que está sendo visível.</span><span class="sxs-lookup"><span data-stu-id="ef3de-244">In the following example, the button that is enabled is on the very same contextual tab that is being made visible.</span></span>

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

## <a name="localizing-the-json-blob"></a><span data-ttu-id="ef3de-245">Localização do blob JSON</span><span class="sxs-lookup"><span data-stu-id="ef3de-245">Localizing the JSON blob</span></span>

<span data-ttu-id="ef3de-246">A bolha JSON que é passada `requestCreateControls` não é localizada da mesma forma que a marcação manifesto para guias de núcleo personalizadas é localizada (que é descrita na localização do Controle a partir do [manifesto](../develop/localization.md#control-localization-from-the-manifest)).</span><span class="sxs-lookup"><span data-stu-id="ef3de-246">The JSON blob that is passed to `requestCreateControls` is not localized the same way that the manifest markup for custom core tabs is localized (which is described at [Control localization from the manifest](../develop/localization.md#control-localization-from-the-manifest)).</span></span> <span data-ttu-id="ef3de-247">Em vez disso, a localização deve ocorrer no tempo de execução usando bolhas JSON distintas para cada localidade.</span><span class="sxs-lookup"><span data-stu-id="ef3de-247">Instead, the localization must occur at runtime using distinct JSON blobs for each locale.</span></span> <span data-ttu-id="ef3de-248">Sugerimos que você use uma `switch` instrução que testa a propriedade [Office.context.displayLanguage.](/javascript/api/office/office.context#displayLanguage)</span><span class="sxs-lookup"><span data-stu-id="ef3de-248">We suggest that you use a `switch` statement that tests the [Office.context.displayLanguage](/javascript/api/office/office.context#displayLanguage) property.</span></span> <span data-ttu-id="ef3de-249">Veja um exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="ef3de-249">The following is an example:</span></span>

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

<span data-ttu-id="ef3de-250">Em seguida, seu código chama a função para obter a bolha localizada que é passada para `requestCreateControls` , como no exemplo a seguir:</span><span class="sxs-lookup"><span data-stu-id="ef3de-250">Then your code calls the function to get the localized blob that is passed to `requestCreateControls`, as in the following example:</span></span>

```javascript
var contextualTabJSON = GetContextualTabsJsonSupportedLocale();
```

## <a name="best-practices-for-custom-contextual-tabs"></a><span data-ttu-id="ef3de-251">Melhores práticas para guias contextuais personalizadas</span><span class="sxs-lookup"><span data-stu-id="ef3de-251">Best practices for custom contextual tabs</span></span>

### <a name="implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported"></a><span data-ttu-id="ef3de-252">Implemente uma experiência de interface do usuário alternativa quando as guias contextuais personalizadas não forem suportadas</span><span class="sxs-lookup"><span data-stu-id="ef3de-252">Implement an alternate UI experience when custom contextual tabs are not supported</span></span>

<span data-ttu-id="ef3de-253">Algumas combinações de plataforma, Office aplicativo e Office build não suportam `requestCreateControls` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-253">Some combinations of platform, Office application, and Office build don't support `requestCreateControls`.</span></span> <span data-ttu-id="ef3de-254">Seu complemento deve ser projetado para fornecer uma experiência alternativa aos usuários que estão executando o complemento em uma dessas combinações.</span><span class="sxs-lookup"><span data-stu-id="ef3de-254">Your add-in should be designed to provide an alternate experience to users who are running the add-in on one of those combinations.</span></span> <span data-ttu-id="ef3de-255">As seções a seguir descrevem duas maneiras de fornecer uma experiência de recuo.</span><span class="sxs-lookup"><span data-stu-id="ef3de-255">The following sections describe two ways of providing a fallback experience.</span></span>

#### <a name="use-noncontextual-tabs-or-controls"></a><span data-ttu-id="ef3de-256">Use guias ou controles não contextuais</span><span class="sxs-lookup"><span data-stu-id="ef3de-256">Use noncontextual tabs or controls</span></span>

<span data-ttu-id="ef3de-257">Há um elemento manifesto, [OverriddenByRibbonApi,](../reference/manifest/overriddenbyribbonapi.md)que foi projetado para criar uma experiência de recuo em um complemento que implementa guias contextuais personalizadas quando o complemento está sendo executado em um aplicativo ou plataforma que não suporta guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-257">There is a manifest element, [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md), that is designed to create a fallback experience in an add-in that implements custom contextual tabs when the add-in is running on an application or platform that doesn't support custom contextual tabs.</span></span> 

<span data-ttu-id="ef3de-258">A estratégia mais simples para usar este elemento é que você define nas guias de núcleo manifesto ou mais personalizadas (ou seja, guias personalizadas *não contextuais)* que duplicam as personalizações de fita das guias contextuais personalizadas em seu complemento.</span><span class="sxs-lookup"><span data-stu-id="ef3de-258">The simplest strategy for using this element is that you define in the manifest one or more custom core tabs (that is, *noncontextual* custom tabs) that duplicate the ribbon customizations of the custom contextual tabs in your add-in.</span></span> <span data-ttu-id="ef3de-259">Mas você adiciona `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` como o elemento primeiro filho do [CustomTab](../reference/manifest/customtab.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-259">But you add `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` as the first child element of the [CustomTab](../reference/manifest/customtab.md).</span></span> <span data-ttu-id="ef3de-260">O efeito de fazê-lo é o seguinte:</span><span class="sxs-lookup"><span data-stu-id="ef3de-260">The effect of doing so is the following:</span></span>

- <span data-ttu-id="ef3de-261">Se o complemento for executado em um aplicativo e plataforma que suportem guias contextuais personalizadas, a guia central personalizada não aparecerá na fita.</span><span class="sxs-lookup"><span data-stu-id="ef3de-261">If the add-in runs on an application and platform that support custom contextual tabs, then the custom core tab won't appear on the ribbon.</span></span> <span data-ttu-id="ef3de-262">Em vez disso, a guia contextual personalizada será criada quando o complemento chamar o `requestCreateControls` método.</span><span class="sxs-lookup"><span data-stu-id="ef3de-262">Instead, the custom contextual tab will be created when the add-in calls the `requestCreateControls` method.</span></span>
- <span data-ttu-id="ef3de-263">Se o complemento for executado em um aplicativo ou plataforma que *não* `requestCreateControls` suporte, a guia central personalizada aparecerá na fita.</span><span class="sxs-lookup"><span data-stu-id="ef3de-263">If the add-in runs on an application or platform that *doesn't* support `requestCreateControls`, then the custom core tab does appear on the ribbon.</span></span>

<span data-ttu-id="ef3de-264">A seguir, um exemplo dessa simples estratégia.</span><span class="sxs-lookup"><span data-stu-id="ef3de-264">The following is an example of this simple strategy.</span></span>

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

<span data-ttu-id="ef3de-265">Esta estratégia simples usa uma guia central personalizada que espelha uma guia contextual personalizada com seus grupos e controles infantis, mas você pode usar uma estratégia mais complexa.</span><span class="sxs-lookup"><span data-stu-id="ef3de-265">This simple strategy uses a custom core tab that mirrors a custom contextual tab with it's child groups and controls, but you can use a more complex strategy.</span></span> <span data-ttu-id="ef3de-266">O `<OverriddenByRibbonApi>` elemento também pode ser adicionado como (o primeiro) elemento infantil aos elementos [grupo](../reference/manifest/group.md) e [controle](../reference/manifest/control.md) (tanto tipo [de botão](../reference/manifest/control.md#button-control) quanto tipo de [menu](../reference/manifest/control.md#menu-dropdown-button-controls)), e elementos do `<Item>` menu.</span><span class="sxs-lookup"><span data-stu-id="ef3de-266">The `<OverriddenByRibbonApi>` element can also be added as (the first) child element to the [Group](../reference/manifest/group.md) and [Control](../reference/manifest/control.md) elements (both [button type](../reference/manifest/control.md#button-control) and [menu type](../reference/manifest/control.md#menu-dropdown-button-controls)), and menu `<Item>` elements.</span></span> <span data-ttu-id="ef3de-267">Este fato permite que você distribua os grupos e controles que de outra forma apareceriam na guia contextual entre vários grupos, botões e menus em várias guias de núcleo personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-267">This fact enables you to distribute the groups and controls that would otherwise appear on the contextual tab among various groups, buttons, and menus in various custom core tabs.</span></span> <span data-ttu-id="ef3de-268">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="ef3de-268">The following is an example.</span></span> <span data-ttu-id="ef3de-269">Observe que "MyButton" aparecerá na guia principal personalizada somente quando as guias contextuais personalizadas não forem suportadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-269">Note that "MyButton" will appear on the custom core tab only when custom contextual tabs are not supported.</span></span> <span data-ttu-id="ef3de-270">Mas o grupo pai e a guia central personalizada aparecerão independentemente de as guias contextuais personalizadas serem suportadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-270">But the parent group and custom core tab will appear regardless of whether custom contextual tabs are supported.</span></span>

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

<span data-ttu-id="ef3de-271">Para obter mais exemplos, consulte [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-271">For more examples, see [OverriddenByRibbonApi](../reference/manifest/overriddenbyribbonapi.md).</span></span>

<span data-ttu-id="ef3de-272">Quando uma guia, grupo ou menu dos pais é marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` com , então ele não é visível, e toda a marcação de criança é ignorada, quando as guias contextuais personalizadas não são suportadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-272">When a parent tab, group, or menu is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, then it isn't visible, and all of it's child markup is ignored, when custom contextual tabs aren't supported.</span></span> <span data-ttu-id="ef3de-273">Então, não importa se algum desses elementos infantis tem o `<OverriddenByRibbonApi>` elemento ou qual é o seu valor.</span><span class="sxs-lookup"><span data-stu-id="ef3de-273">So, it doesn't matter if any of those child elements have the `<OverriddenByRibbonApi>` element or what its value is.</span></span> <span data-ttu-id="ef3de-274">A implicação disso é que se um item, controle ou grupo do menu deve ser visível em todos os contextos, então não só não deve ser marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` com , mas seu *menu, grupo e guia ancestrais também não devem ser marcados dessa forma*.</span><span class="sxs-lookup"><span data-stu-id="ef3de-274">The implication of this is that if a menu item, control, or group must be visible in all contexts, then not only should it not be marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`, but *its ancestor menu, group, and tab must also not be marked this way*.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef3de-275">Não marque *todos os* elementos infantis de uma guia, grupo ou menu com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-275">Don't mark *all* of the child elements of a tab, group, or menu with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span> <span data-ttu-id="ef3de-276">Isso é inútil se o elemento pai estiver marcado `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` por razões dadas no parágrafo anterior.</span><span class="sxs-lookup"><span data-stu-id="ef3de-276">This is pointless if the parent element is marked with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` for reasons given in the preceding paragraph.</span></span> <span data-ttu-id="ef3de-277">Além disso, se você deixar de fora `<OverriddenByRibbonApi>` o pai (ou defini-lo `false` para ), então o pai aparecerá independentemente de as guias contextuais personalizadas serem suportadas, mas estará vazia quando elas forem suportadas.</span><span class="sxs-lookup"><span data-stu-id="ef3de-277">Moreover, if you leave out the `<OverriddenByRibbonApi>` on the parent (or set it to `false`), then the parent will appear regardless of whether custom contextual tabs are supported, but it will be empty when they are supported.</span></span> <span data-ttu-id="ef3de-278">Assim, se todos os elementos da criança não aparecerem quando as guias contextuais personalizadas forem suportadas, marque o pai e apenas o pai, com `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>` .</span><span class="sxs-lookup"><span data-stu-id="ef3de-278">So, if all the child elements shouldn't appear when custom contextual tabs are supported, mark the parent, and only the parent, with `<OverriddenByRibbonApi>true</OverriddenByRibbonApi>`.</span></span>

#### <a name="use-apis-that-show-or-hide-a-task-pane-in-specified-contexts"></a><span data-ttu-id="ef3de-279">Use APIs que mostram ou ocultam um painel de tarefas em contextos especificados</span><span class="sxs-lookup"><span data-stu-id="ef3de-279">Use APIs that show or hide a task pane in specified contexts</span></span>

<span data-ttu-id="ef3de-280">Como alternativa, `<OverriddenByRibbonApi>` seu complemento pode definir um painel de tarefas com controles de interface do usuário que duplicam a funcionalidade dos controles em uma guia contextual personalizada. Em seguida, use os métodos [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) e [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) para mostrar o painel de tarefas quando, e somente quando, a guia contextual teria sido mostrada se fosse suportada.</span><span class="sxs-lookup"><span data-stu-id="ef3de-280">As an alternative to `<OverriddenByRibbonApi>`, your add-in can define a task pane with UI controls that duplicate the functionality of the controls on a custom contextual tab. Then use the [Office.addin.showAsTaskpane](/javascript/api/office/office.addin?view=common-js&preserve-view=true#showAsTaskpane__) and [Office.addin.hide](/javascript/api/office/office.addin?view=common-js&preserve-view=true#hide__) methods to show the task pane when, and only when, the contextual tab would have been shown if it was supported.</span></span> <span data-ttu-id="ef3de-281">Para obter detalhes sobre como usar esses métodos, consulte [Mostrar ou ocultar o painel de tarefas do seu Office Add-in](../develop/show-hide-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="ef3de-281">For details on how to use these methods, see [Show or hide the task pane of your Office Add-in](../develop/show-hide-add-in.md).</span></span>

### <a name="handle-the-hostrestartneeded-error"></a><span data-ttu-id="ef3de-282">Manuseie o erro HostRestartNeed</span><span class="sxs-lookup"><span data-stu-id="ef3de-282">Handle the HostRestartNeeded error</span></span>

<span data-ttu-id="ef3de-283">Em alguns cenários, o Office não consegue atualizar a faixa de opções e retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="ef3de-283">In some scenarios, Office is unable to update the ribbon and will return an error.</span></span> <span data-ttu-id="ef3de-284">Por exemplo, se o suplemento for atualizado e o suplemento atualizado tiver um conjunto diferente de comandos de suplemento personalizados, o aplicativo do Office deverá ser fechado e reaberto.</span><span class="sxs-lookup"><span data-stu-id="ef3de-284">For example, if the add-in is upgraded and the upgraded add-in has a different set of custom add-in commands, then the Office application must be closed and reopened.</span></span> <span data-ttu-id="ef3de-285">Até que isso ocorra, o método `requestUpdate` retornará o erro `HostRestartNeeded`.</span><span class="sxs-lookup"><span data-stu-id="ef3de-285">Until it is, the `requestUpdate` method will return the error `HostRestartNeeded`.</span></span> <span data-ttu-id="ef3de-286">Seu código deve lidar com esse erro.</span><span class="sxs-lookup"><span data-stu-id="ef3de-286">Your code should handle this error.</span></span> <span data-ttu-id="ef3de-287">A seguir, um exemplo de como.</span><span class="sxs-lookup"><span data-stu-id="ef3de-287">The following is an example of how.</span></span> <span data-ttu-id="ef3de-288">Nesse caso, o método `reportError` exibe o erro para o usuário.</span><span class="sxs-lookup"><span data-stu-id="ef3de-288">In this case, the `reportError` method displays the error to the user.</span></span>

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
