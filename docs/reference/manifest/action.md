---
title: Elemento Action no arquivo de manifesto
description: Este elemento Especifica a ação a ser executada quando o usuário seleciona um botão ou controle de menu.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: e345d0a1682e0125373a309e1e56eb2d6298ac7d
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771407"
---
# <a name="action-element"></a><span data-ttu-id="330ad-103">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="330ad-103">Action element</span></span>

<span data-ttu-id="330ad-104">Especifica a ação a ser executada quando o usuário seleciona um controle de  [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) .</span><span class="sxs-lookup"><span data-stu-id="330ad-104">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control.</span></span>

## <a name="attributes"></a><span data-ttu-id="330ad-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="330ad-105">Attributes</span></span>

|  <span data-ttu-id="330ad-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="330ad-106">Attribute</span></span>  |  <span data-ttu-id="330ad-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="330ad-107">Required</span></span>  |  <span data-ttu-id="330ad-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="330ad-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="330ad-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="330ad-109">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="330ad-110">Sim</span><span class="sxs-lookup"><span data-stu-id="330ad-110">Yes</span></span>  | <span data-ttu-id="330ad-111">Tipo de ação a executar</span><span class="sxs-lookup"><span data-stu-id="330ad-111">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="330ad-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="330ad-112">Child elements</span></span>

|  <span data-ttu-id="330ad-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="330ad-113">Element</span></span> |  <span data-ttu-id="330ad-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="330ad-114">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="330ad-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="330ad-115">FunctionName</span></span>](#functionname) |    <span data-ttu-id="330ad-116">Especifica o nome da função a executar.</span><span class="sxs-lookup"><span data-stu-id="330ad-116">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="330ad-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="330ad-117">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="330ad-118">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="330ad-118">Specifies the source file location for this action.</span></span> |
|  [<span data-ttu-id="330ad-119">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="330ad-119">TaskpaneId</span></span>](#taskpaneid) | <span data-ttu-id="330ad-120">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="330ad-120">Specifies the ID of the task pane container.</span></span>|
|  [<span data-ttu-id="330ad-121">Title</span><span class="sxs-lookup"><span data-stu-id="330ad-121">Title</span></span>](#title) | <span data-ttu-id="330ad-122">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="330ad-122">Specifies the custom title for the task pane.</span></span>|
|  [<span data-ttu-id="330ad-123">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="330ad-123">SupportsPinning</span></span>](#supportspinning) | <span data-ttu-id="330ad-124">Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="330ad-124">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="330ad-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="330ad-125">xsi:type</span></span>

<span data-ttu-id="330ad-p101">Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:</span><span class="sxs-lookup"><span data-stu-id="330ad-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="330ad-128">FunctionName</span><span class="sxs-lookup"><span data-stu-id="330ad-128">FunctionName</span></span>

<span data-ttu-id="330ad-p102">Elemento obrigatório quando **xsi:type** é "ExecuteFunction". Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="330ad-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="330ad-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="330ad-132">SourceLocation</span></span>

<span data-ttu-id="330ad-133">Elemento obrigatório quando **xsi: Type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="330ad-133">Required element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="330ad-134">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="330ad-134">Specifies the source file location for this action.</span></span> <span data-ttu-id="330ad-135">O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **URL** no elemento **URLs** do elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="330ad-135">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="330ad-136">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="330ad-136">TaskpaneId</span></span>

<span data-ttu-id="330ad-p104">Elemento opcional quando **xsi:type** for "ShowTaskpane". Especifica a ID do contêiner do painel de tarefas. Quando você tiver várias ações "ShowTaskpane", use uma **TaskpaneId** diferente se desejar ter um painel independente para cada uma. Use a mesma **TaskpaneId** para diferentes ações que compartilhem o mesmo painel. Quando os usuários escolhem comandos que compartilham o mesmo **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação correspondente "SourceLocation".</span><span class="sxs-lookup"><span data-stu-id="330ad-p104">Optional element when  **xsi:type** is "ShowTaskpane". Specifies the ID of the task pane container. When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each. Use the same **TaskpaneId** for  different actions that share the same pane. When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="330ad-142">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="330ad-142">This element is not supported in Outlook.</span></span>

<span data-ttu-id="330ad-143">O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="330ad-143">The following example shows two actions that share the same **TaskpaneId**.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="aTaskPaneUrl" />
</Action>

<Action xsi:type="ShowTaskpane">
  <TaskpaneId>MyPane</TaskpaneId>
  <SourceLocation resid="anotherTaskPaneUrl" />
</Action>
```  

<span data-ttu-id="330ad-p105">O exemplo a seguir mostra duas ações que usam um **TaskpaneId** diferente. Para ver esses exemplos em contexto, consulte [Exemplo de comando de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="330ad-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID1</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane1.Url" />
</Action>

<Action xsi:type="ShowTaskpane">
   <TaskpaneId>MyTaskPaneID2</TaskpaneId>
   <SourceLocation resid="Contoso.Taskpane2.Url" />
</Action>
```  

```xml
<bt:Urls>
   <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
   <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
</bt:Urls>
```  

## <a name="title"></a><span data-ttu-id="330ad-146">Cargo</span><span class="sxs-lookup"><span data-stu-id="330ad-146">Title</span></span>

<span data-ttu-id="330ad-147">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="330ad-147">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="330ad-148">Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="330ad-148">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="330ad-149">O exemplo a seguir mostra uma ação que usa o elemento **title** .</span><span class="sxs-lookup"><span data-stu-id="330ad-149">The following example shows an action that uses the **Title** element.</span></span> <span data-ttu-id="330ad-150">Observe que você não atribui o **título** a uma cadeia de caracteres diretamente.</span><span class="sxs-lookup"><span data-stu-id="330ad-150">Note that you don't assign the **Title** to a string directly.</span></span> <span data-ttu-id="330ad-151">Em vez disso, você atribui a ele uma ID de recurso (Resid), que é definida na seção **recursos** do manifesto e não pode ter mais de 32 caracteres.</span><span class="sxs-lookup"><span data-stu-id="330ad-151">Instead, you assign it a resource ID (resid), that is defined in the **Resources** section of the manifest and can be no more than 32 characters.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="PG.Code.Url" />
    <Title resid="PG.CodeCommand.Title" />
</Action>

 ... Other markup omitted ...
<Resources>
    <bt:Images> ...
    </bt:Images>
    <bt:Urls>
        <bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
    </bt:ShortStrings>
 ... Other markup omitted ...
</Resources>
```

## <a name="supportspinning"></a><span data-ttu-id="330ad-152">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="330ad-152">SupportsPinning</span></span>

<span data-ttu-id="330ad-153">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="330ad-153">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="330ad-154">Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="330ad-154">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="330ad-155">Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="330ad-155">Include this element with a value of `true` to support task pane pinning.</span></span> <span data-ttu-id="330ad-156">O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção.</span><span class="sxs-lookup"><span data-stu-id="330ad-156">The user will be able to "pin" the task pane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="330ad-157">Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="330ad-157">For more information, see [Implement a pinnable task pane in Outlook](../../outlook/pinnable-taskpane.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="330ad-158">Embora o `SupportsPinning` elemento tenha sido introduzido no [conjunto de requisitos 1,5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), atualmente só há suporte para assinantes do Microsoft 365 usando o seguinte.</span><span class="sxs-lookup"><span data-stu-id="330ad-158">Although the `SupportsPinning` element was introduced in [requirement set 1.5](../objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only supported for Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="330ad-159">Outlook 2016 ou posterior no Windows (compilação 7628,1000 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="330ad-159">Outlook 2016 or later on Windows (build 7628.1000 or later)</span></span>
> - <span data-ttu-id="330ad-160">Outlook 2016 ou posterior no Mac (Build 16.13.503 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="330ad-160">Outlook 2016 or later on Mac (build 16.13.503 or later)</span></span>
> - <span data-ttu-id="330ad-161">Outlook na Web moderno</span><span class="sxs-lookup"><span data-stu-id="330ad-161">Modern Outlook on the web</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
