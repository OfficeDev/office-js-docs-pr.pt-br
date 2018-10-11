
# <a name="context"></a><span data-ttu-id="71b1a-101">context</span><span class="sxs-lookup"><span data-stu-id="71b1a-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="71b1a-102">.context do [Office](Office.md)</span><span class="sxs-lookup"><span data-stu-id="71b1a-102">[Office](Office.md).context</span></span>

<span data-ttu-id="71b1a-p101">O namespace do Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="71b1a-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="71b1a-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71b1a-105">Requirements</span></span>

|<span data-ttu-id="71b1a-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="71b1a-106">Requirement</span></span>| <span data-ttu-id="71b1a-107">Valor</span><span class="sxs-lookup"><span data-stu-id="71b1a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b1a-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71b1a-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b1a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="71b1a-109">1.0</span></span>|
|[<span data-ttu-id="71b1a-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71b1a-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71b1a-111">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="71b1a-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="71b1a-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="71b1a-112">Namespaces</span></span>

<span data-ttu-id="71b1a-113">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="71b1a-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="71b1a-114">Membros</span><span class="sxs-lookup"><span data-stu-id="71b1a-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="71b1a-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="71b1a-115">displayLanguage :String</span></span>

<span data-ttu-id="71b1a-116">Obtém a localidade (idioma) no formato de marca de linguagem RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="71b1a-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="71b1a-117">O valor `displayLanguage` reflete a configuração atual do **Idioma de Exibição** especificada em **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="71b1a-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="71b1a-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="71b1a-118">Type:</span></span>

*   <span data-ttu-id="71b1a-119">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="71b1a-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="71b1a-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71b1a-120">Requirements</span></span>

|<span data-ttu-id="71b1a-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="71b1a-121">Requirement</span></span>| <span data-ttu-id="71b1a-122">Valor</span><span class="sxs-lookup"><span data-stu-id="71b1a-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b1a-123">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71b1a-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b1a-124">1.0</span><span class="sxs-lookup"><span data-stu-id="71b1a-124">1.0</span></span>|
|[<span data-ttu-id="71b1a-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71b1a-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71b1a-126">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="71b1a-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b1a-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="71b1a-127">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="71b1a-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="71b1a-128">officeTheme :Object</span></span>

<span data-ttu-id="71b1a-129">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="71b1a-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="71b1a-130">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="71b1a-130">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="71b1a-p102">Usando as cores de tema do Office, você pode coordenar o esquema de cores do seu suplemento com o tema atual do Office, selecionado pelo usuário em **Arquivo > Conta do Office > Tema da interface de usuário do Office**, que é aplicado a todos os aplicativos host do Office. Usar cores de tema do Office é apropriado para suplementos de painel de tarefas e email.</span><span class="sxs-lookup"><span data-stu-id="71b1a-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="71b1a-133">Tipo:</span><span class="sxs-lookup"><span data-stu-id="71b1a-133">Type:</span></span>

*   <span data-ttu-id="71b1a-134">Objeto</span><span class="sxs-lookup"><span data-stu-id="71b1a-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="71b1a-135">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="71b1a-135">Properties:</span></span>

|<span data-ttu-id="71b1a-136">Nome</span><span class="sxs-lookup"><span data-stu-id="71b1a-136">Name</span></span>| <span data-ttu-id="71b1a-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="71b1a-137">Type</span></span>| <span data-ttu-id="71b1a-138">Descrição</span><span class="sxs-lookup"><span data-stu-id="71b1a-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="71b1a-139">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="71b1a-139">String</span></span>|<span data-ttu-id="71b1a-140">Obtém a cor do plano de fundo do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="71b1a-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="71b1a-141">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="71b1a-141">String</span></span>|<span data-ttu-id="71b1a-142">Obtém a cor de primeiro plano do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="71b1a-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="71b1a-143">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="71b1a-143">String</span></span>|<span data-ttu-id="71b1a-144">Obtém o tema do Office para controlar a cor do plano de fundo como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="71b1a-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="71b1a-145">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="71b1a-145">String</span></span>|<span data-ttu-id="71b1a-146">Obtém a cor de controle do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="71b1a-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71b1a-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71b1a-147">Requirements</span></span>

|<span data-ttu-id="71b1a-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="71b1a-148">Requirement</span></span>| <span data-ttu-id="71b1a-149">Valor</span><span class="sxs-lookup"><span data-stu-id="71b1a-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b1a-150">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71b1a-150">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b1a-151">1.3</span><span class="sxs-lookup"><span data-stu-id="71b1a-151">1.3</span></span>|
|[<span data-ttu-id="71b1a-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71b1a-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71b1a-153">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="71b1a-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="71b1a-154">Exemplo</span><span class="sxs-lookup"><span data-stu-id="71b1a-154">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="71b1a-155">roamingSettings:[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="71b1a-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="71b1a-156">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email salvo na caixa de correio de um usuário.</span><span class="sxs-lookup"><span data-stu-id="71b1a-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="71b1a-157">O objeto `RoamingSettings` permite armazenar e acessar dados para um suplemento de email armazenado na caixa de correio de um usuário, para que ele esteja disponível para esse complemento quando estiver sendo executado em qualquer aplicativo cliente host usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="71b1a-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="71b1a-158">Tipo:</span><span class="sxs-lookup"><span data-stu-id="71b1a-158">Type:</span></span>

*   [<span data-ttu-id="71b1a-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="71b1a-159">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="71b1a-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71b1a-160">Requirements</span></span>

|<span data-ttu-id="71b1a-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="71b1a-161">Requirement</span></span>| <span data-ttu-id="71b1a-162">Valor</span><span class="sxs-lookup"><span data-stu-id="71b1a-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="71b1a-163">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71b1a-163">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="71b1a-164">1.0</span><span class="sxs-lookup"><span data-stu-id="71b1a-164">1.0</span></span>|
|[<span data-ttu-id="71b1a-165">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="71b1a-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="71b1a-166">Restrito</span><span class="sxs-lookup"><span data-stu-id="71b1a-166">Restricted</span></span>|
|[<span data-ttu-id="71b1a-167">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71b1a-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="71b1a-168">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="71b1a-168">Compose or read</span></span>|