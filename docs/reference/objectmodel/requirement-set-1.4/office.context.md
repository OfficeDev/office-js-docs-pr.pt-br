
# <a name="context"></a><span data-ttu-id="35d9f-101">context</span><span class="sxs-lookup"><span data-stu-id="35d9f-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="35d9f-102">.context do [Office](Office.md)</span><span class="sxs-lookup"><span data-stu-id="35d9f-102">[Office](Office.md).context</span></span>

<span data-ttu-id="35d9f-p101">O namespace do Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="35d9f-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="35d9f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35d9f-105">Requirements</span></span>

|<span data-ttu-id="35d9f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="35d9f-106">Requirement</span></span>| <span data-ttu-id="35d9f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="35d9f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="35d9f-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35d9f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35d9f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="35d9f-109">1.0</span></span>|
|[<span data-ttu-id="35d9f-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35d9f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35d9f-111">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="35d9f-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="35d9f-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="35d9f-112">Namespaces</span></span>

<span data-ttu-id="35d9f-113">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="35d9f-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="35d9f-114">Membros</span><span class="sxs-lookup"><span data-stu-id="35d9f-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="35d9f-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="35d9f-115">displayLanguage :String</span></span>

<span data-ttu-id="35d9f-116">Obtém a localidade (idioma) no formato de marca de linguagem RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="35d9f-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="35d9f-117">O valor `displayLanguage` reflete a configuração atual do **Idioma de Exibição** especificada em **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="35d9f-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="35d9f-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="35d9f-118">Type:</span></span>

*   <span data-ttu-id="35d9f-119">String</span><span class="sxs-lookup"><span data-stu-id="35d9f-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35d9f-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35d9f-120">Requirements</span></span>

|<span data-ttu-id="35d9f-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="35d9f-121">Requirement</span></span>| <span data-ttu-id="35d9f-122">Valor</span><span class="sxs-lookup"><span data-stu-id="35d9f-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="35d9f-123">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35d9f-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35d9f-124">1.0</span><span class="sxs-lookup"><span data-stu-id="35d9f-124">1.0</span></span>|
|[<span data-ttu-id="35d9f-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35d9f-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35d9f-126">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="35d9f-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35d9f-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35d9f-127">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="35d9f-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="35d9f-128">officeTheme :Object</span></span>

<span data-ttu-id="35d9f-129">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="35d9f-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="35d9f-130">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35d9f-130">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35d9f-p102">Usando as cores de tema do Office, você pode coordenar o esquema de cores do seu suplemento com o tema atual do Office, selecionado pelo usuário em **Arquivo > Conta do Office > Tema da interface de usuário do Office**, que é aplicado a todos os aplicativos host do Office. Usar cores de tema do Office é apropriado para suplementos de painel de tarefas e email.</span><span class="sxs-lookup"><span data-stu-id="35d9f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="35d9f-133">Tipo:</span><span class="sxs-lookup"><span data-stu-id="35d9f-133">Type:</span></span>

*   <span data-ttu-id="35d9f-134">Objeto</span><span class="sxs-lookup"><span data-stu-id="35d9f-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="35d9f-135">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="35d9f-135">Properties:</span></span>

|<span data-ttu-id="35d9f-136">Nome</span><span class="sxs-lookup"><span data-stu-id="35d9f-136">Name</span></span>| <span data-ttu-id="35d9f-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="35d9f-137">Type</span></span>| <span data-ttu-id="35d9f-138">Descrição</span><span class="sxs-lookup"><span data-stu-id="35d9f-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="35d9f-139">String</span><span class="sxs-lookup"><span data-stu-id="35d9f-139">String</span></span>|<span data-ttu-id="35d9f-140">Obtém a cor do plano de fundo do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="35d9f-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="35d9f-141">String</span><span class="sxs-lookup"><span data-stu-id="35d9f-141">String</span></span>|<span data-ttu-id="35d9f-142">Obtém a cor de primeiro plano do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="35d9f-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="35d9f-143">String</span><span class="sxs-lookup"><span data-stu-id="35d9f-143">String</span></span>|<span data-ttu-id="35d9f-144">Obtém o tema do Office para controlar a cor do plano de fundo como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="35d9f-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="35d9f-145">String</span><span class="sxs-lookup"><span data-stu-id="35d9f-145">String</span></span>|<span data-ttu-id="35d9f-146">Obtém a cor de controle do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="35d9f-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35d9f-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35d9f-147">Requirements</span></span>

|<span data-ttu-id="35d9f-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="35d9f-148">Requirement</span></span>| <span data-ttu-id="35d9f-149">Valor</span><span class="sxs-lookup"><span data-stu-id="35d9f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="35d9f-150">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35d9f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35d9f-151">1.3</span><span class="sxs-lookup"><span data-stu-id="35d9f-151">1.3</span></span>|
|[<span data-ttu-id="35d9f-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35d9f-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35d9f-153">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="35d9f-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="35d9f-154">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35d9f-154">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="35d9f-155">roamingSettings:[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="35d9f-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="35d9f-156">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email salvo na caixa de correio de um usuário.</span><span class="sxs-lookup"><span data-stu-id="35d9f-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="35d9f-157">O objeto `RoamingSettings` permite armazenar e acessar dados para um suplemento de email armazenado na caixa de correio de um usuário, para que ele esteja disponível para esse complemento quando estiver sendo executado em qualquer aplicativo cliente host usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="35d9f-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="35d9f-158">Tipo:</span><span class="sxs-lookup"><span data-stu-id="35d9f-158">Type:</span></span>

*   [<span data-ttu-id="35d9f-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="35d9f-159">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="35d9f-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35d9f-160">Requirements</span></span>

|<span data-ttu-id="35d9f-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="35d9f-161">Requirement</span></span>| <span data-ttu-id="35d9f-162">Valor</span><span class="sxs-lookup"><span data-stu-id="35d9f-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="35d9f-163">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35d9f-163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35d9f-164">1.0</span><span class="sxs-lookup"><span data-stu-id="35d9f-164">1.0</span></span>|
|[<span data-ttu-id="35d9f-165">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35d9f-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35d9f-166">Restrito</span><span class="sxs-lookup"><span data-stu-id="35d9f-166">Restricted</span></span>|
|[<span data-ttu-id="35d9f-167">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35d9f-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35d9f-168">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="35d9f-168">Compose or read</span></span>|