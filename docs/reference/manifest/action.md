# <a name="action-element"></a><span data-ttu-id="231d0-101">Elemento Action</span><span class="sxs-lookup"><span data-stu-id="231d0-101">Action element</span></span>

<span data-ttu-id="231d0-102">Especifica a ação a ser executada quando o usuário seleciona controles de [Button](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="231d0-102">Specifies the action to perform when the user selects a  [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="231d0-103">Atributos</span><span class="sxs-lookup"><span data-stu-id="231d0-103">Attributes</span></span>

|  <span data-ttu-id="231d0-104">Atributo</span><span class="sxs-lookup"><span data-stu-id="231d0-104">Attribute</span></span>  |  <span data-ttu-id="231d0-105">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="231d0-105">Required</span></span>  |  <span data-ttu-id="231d0-106">Descrição</span><span class="sxs-lookup"><span data-stu-id="231d0-106">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="231d0-107">xsi:type</span><span class="sxs-lookup"><span data-stu-id="231d0-107">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="231d0-108">Sim</span><span class="sxs-lookup"><span data-stu-id="231d0-108">Yes</span></span>  | <span data-ttu-id="231d0-109">Tipo de ação a executar</span><span class="sxs-lookup"><span data-stu-id="231d0-109">Action type to take</span></span>|

## <a name="child-elements"></a><span data-ttu-id="231d0-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="231d0-110">Child elements</span></span>

|  <span data-ttu-id="231d0-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="231d0-111">Element</span></span> |  <span data-ttu-id="231d0-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="231d0-112">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="231d0-113">FunctionName</span><span class="sxs-lookup"><span data-stu-id="231d0-113">FunctionName</span></span>](#functionname) |    <span data-ttu-id="231d0-114">Especifica o nome da função a executar.</span><span class="sxs-lookup"><span data-stu-id="231d0-114">Specifies the name of the function to execute.</span></span> |
|  [<span data-ttu-id="231d0-115">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="231d0-115">SourceLocation</span></span>](#sourcelocation) |    <span data-ttu-id="231d0-116">Especifica o local do arquivo de origem para essa ação.</span><span class="sxs-lookup"><span data-stu-id="231d0-116">Specifies the source file location for this action.</span></span> |
| <span data-ttu-id="231d0-117"> [TaskpaneId](#taskpaneid)</span><span class="sxs-lookup"><span data-stu-id="231d0-117"> [TaskpaneId](#taskpaneid)</span></span> | <span data-ttu-id="231d0-118">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="231d0-118">Specifies the ID of the task pane container.</span></span>|
| <span data-ttu-id="231d0-119"> [Title](#title)</span><span class="sxs-lookup"><span data-stu-id="231d0-119"> [Title](#title)</span></span> | <span data-ttu-id="231d0-120">Especifica o título personalizado do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="231d0-120">Specifies the custom title for the task pane.</span></span>|
| <span data-ttu-id="231d0-121"> [SupportsPinning](#supportspinning)</span><span class="sxs-lookup"><span data-stu-id="231d0-121"> [SupportsPinning](#supportspinning)</span></span> | <span data-ttu-id="231d0-122">Especifica se um painel de tarefas tem suporte para fixação, que mantém o painel de tarefas aberto quando o usuário altera a seleção.</span><span class="sxs-lookup"><span data-stu-id="231d0-122">Specifies that a task pane supports pinning, which keeps the task pane open when the user changes the selection.</span></span>|
  

## <a name="xsitype"></a><span data-ttu-id="231d0-123">xsi:type</span><span class="sxs-lookup"><span data-stu-id="231d0-123">xsi:type</span></span>

<span data-ttu-id="231d0-p101">Este atributo especifica o tipo de ação realizada quando o usuário seleciona o botão. Pode ser uma das seguintes:</span><span class="sxs-lookup"><span data-stu-id="231d0-p101">This attribute specifies the kind of action performed when the user selects the button. It can be one of the following:</span></span>

- `ExecuteFunction`
- `ShowTaskpane`

## <a name="functionname"></a><span data-ttu-id="231d0-126">FunctionName</span><span class="sxs-lookup"><span data-stu-id="231d0-126">FunctionName</span></span>

<span data-ttu-id="231d0-p102">Elemento obrigatório quando **xsi:type** é "ExecuteFunction". Especifica o nome da função a ser executada. A função está contida no arquivo especificado no elemento [FunctionFile](functionfile.md).</span><span class="sxs-lookup"><span data-stu-id="231d0-p102">Required element when **xsi:type** is "ExecuteFunction". Specifies the name of the function to execute. The function is contained in the file specified in the [FunctionFile](functionfile.md) element.</span></span>

```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>getSubject</FunctionName>
</Action>
```

## <a name="sourcelocation"></a><span data-ttu-id="231d0-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="231d0-130">SourceLocation</span></span>

<span data-ttu-id="231d0-p103">Elemento obrigatório quando **xsi:type** for "ShowTaskpane". Especifica o local do arquivo de origem para essa ação. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="231d0-p103">Required element when  **xsi:type** is "ShowTaskpane". Specifies the source file location for this action. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the [Resources](resources.md) element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
</Action>
```  

## <a name="taskpaneid"></a><span data-ttu-id="231d0-134">TaskpaneId</span><span class="sxs-lookup"><span data-stu-id="231d0-134">TaskpaneId</span></span>

<span data-ttu-id="231d0-135">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="231d0-135">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="231d0-136">Especifica a ID do contêiner do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="231d0-136">Specifies the ID of the task pane container.</span></span> <span data-ttu-id="231d0-137">Quando você tem várias ações "ShowTaskpane", use uma **TaskpaneId** diferente, se prefere ter um painel independente para cada uma.</span><span class="sxs-lookup"><span data-stu-id="231d0-137">When you have multiple "ShowTaskpane" actions, use a different **TaskpaneId** if you want an independent pane for each.</span></span> <span data-ttu-id="231d0-138">Use a mesma **TaskpaneId** para diferentes ações que compartilham o mesmo painel.</span><span class="sxs-lookup"><span data-stu-id="231d0-138">Use the same **TaskpaneId** for  different actions that share the same pane.</span></span> <span data-ttu-id="231d0-139">Quando os usuários escolhem comandos que compartilham a mesma **TaskpaneId**, o contêiner do painel permanece aberto, mas o conteúdo do painel é substituído pela ação "SourceLocation" correspondente.</span><span class="sxs-lookup"><span data-stu-id="231d0-139">When users choose commands that share the same **TaskpaneId**, the pane container will remain open but the contents of the pane will be replaced with the corresponding Action "SourceLocation".</span></span>

> [!NOTE]
> <span data-ttu-id="231d0-140">Esse elemento não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="231d0-140">Note: This element is not supported in Outlook.</span></span>

<span data-ttu-id="231d0-141">O exemplo a seguir mostra duas ações que compartilham o mesmo **TaskpaneId**.</span><span class="sxs-lookup"><span data-stu-id="231d0-141">The following example shows two actions that share the same **TaskpaneId**.</span></span>

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

<span data-ttu-id="231d0-p105">O exemplo a seguir mostra duas ações que usam um **TaskpaneId** diferente. Para ver esses exemplos em contexto, consulte [Exemplo de comando de suplemento simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span><span class="sxs-lookup"><span data-stu-id="231d0-p105">The following examples show two actions that use a different **TaskpaneId**. To see these examples in context, see [Simple Add-in Commands Sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/Simple/Manifest/SimpleAddin.xml).</span></span>

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

## <a name="title"></a><span data-ttu-id="231d0-144">Cargo</span><span class="sxs-lookup"><span data-stu-id="231d0-144">Title</span></span>

<span data-ttu-id="231d0-145">Elemento opcional quando  **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="231d0-145">Optional element when  **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="231d0-146">Especifica o título personalizado do painel de tarefas desta ação.</span><span class="sxs-lookup"><span data-stu-id="231d0-146">Specifies the custom title for the task pane for this action.</span></span>

<span data-ttu-id="231d0-147">Os exemplos a seguir mostram duas ações distintas que usam o elemento **Title**.</span><span class="sxs-lookup"><span data-stu-id="231d0-147">The following examples show two different actions that use the **Title** element.</span></span>

```xml
<Action xsi:type="ShowTaskpane">
<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
<SourceLocation resid="PG.Code.Url" />
<Title resid="PG.CodeCommand.Title" />
</Action>
```

```xml
<Action xsi:type="ShowTaskpane">
<SourceLocation resid="PG.Run.Url" />
<Title resid="PG.RunCommand.Title" />
</Action>
```

```xml
<bt:Urls>
<bt:Url id="PG.Code.Url" DefaultValue="https://localhost:3000?commands=1" />
<bt:Url id="PG.Run.Url" DefaultValue="https://localhost:3000/run.html" />
</bt:Urls>
```

```xml
<bt:ShortStrings>
<bt:String id="PG.CodeCommand.Title" DefaultValue="Code" />
<bt:String id="PG.RunCommand.Title" DefaultValue="Run" />
</bt:ShortStrings>
```

## <a name="supportspinning"></a><span data-ttu-id="231d0-148">SupportsPinning</span><span class="sxs-lookup"><span data-stu-id="231d0-148">SupportsPinning</span></span>

<span data-ttu-id="231d0-149">Elemento opcional quando **xsi:type** for "ShowTaskpane".</span><span class="sxs-lookup"><span data-stu-id="231d0-149">Optional element when **xsi:type** is "ShowTaskpane".</span></span> <span data-ttu-id="231d0-150">Os elementos [VersionOverrides](versionoverrides.md) incluídos devem ter um valor `VersionOverridesV1_1` para o atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="231d0-150">The containing [VersionOverrides](versionoverrides.md) elements must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span> <span data-ttu-id="231d0-151">Inclua esse elemento com um valor `true` a fim de fornecer suporte para fixação do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="231d0-151">Include this element with a value of `true` to support taskpane pinning.</span></span> <span data-ttu-id="231d0-152">O usuário pode "fixar" o painel de tarefas, fazendo com que ele permaneça aberto quando alterar a seleção.</span><span class="sxs-lookup"><span data-stu-id="231d0-152">The user will be able to "pin" the taskpane, causing it to stay open when changing the selection.</span></span> <span data-ttu-id="231d0-153">Para saber mais, consulte [Implementar um painel de tarefas fixável no Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="231d0-153">For more information, see [Implement a pinnable taskpane in Outlook](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>

> [!NOTE]
> <span data-ttu-id="231d0-154">Atualmente, o SupportsPinning só é compatível com o Outlook 2016 para Windows (build 7628.1000 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="231d0-154">Note: SupportsPinning currently only supported by Outlook 2016 for Windows (build 7628.1000 or later).</span></span>

```xml
<Action xsi:type="ShowTaskpane">
  <SourceLocation resid="readTaskPaneUrl" />
  <SupportsPinning>true</SupportsPinning>
</Action>
```
