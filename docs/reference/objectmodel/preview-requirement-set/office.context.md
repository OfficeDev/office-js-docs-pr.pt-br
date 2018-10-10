
# <a name="context"></a><span data-ttu-id="3dd79-101">context</span><span class="sxs-lookup"><span data-stu-id="3dd79-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="3dd79-102">.context do [Office](Office.md)</span><span class="sxs-lookup"><span data-stu-id="3dd79-102">[Office](Office.md).context</span></span>

<span data-ttu-id="3dd79-p101">O namespace do Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces usadas pelos suplementos do Outlook. Para obter uma listagem completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="3dd79-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dd79-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dd79-105">Requirements</span></span>

|<span data-ttu-id="3dd79-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dd79-106">Requirement</span></span>| <span data-ttu-id="3dd79-107">Valor</span><span class="sxs-lookup"><span data-stu-id="3dd79-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dd79-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dd79-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dd79-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3dd79-109">1.0</span></span>|
|[<span data-ttu-id="3dd79-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dd79-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dd79-111">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3dd79-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3dd79-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="3dd79-112">Members and methods</span></span>

| <span data-ttu-id="3dd79-113">Membro</span><span class="sxs-lookup"><span data-stu-id="3dd79-113">Member</span></span> | <span data-ttu-id="3dd79-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dd79-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3dd79-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3dd79-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3dd79-116">Membro</span><span class="sxs-lookup"><span data-stu-id="3dd79-116">Member</span></span> |
| [<span data-ttu-id="3dd79-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="3dd79-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="3dd79-118">Membro</span><span class="sxs-lookup"><span data-stu-id="3dd79-118">Member</span></span> |
| [<span data-ttu-id="3dd79-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dd79-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | <span data-ttu-id="3dd79-120">Membro</span><span class="sxs-lookup"><span data-stu-id="3dd79-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3dd79-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="3dd79-121">Namespaces</span></span>

<span data-ttu-id="3dd79-122">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="3dd79-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="3dd79-123">Membros</span><span class="sxs-lookup"><span data-stu-id="3dd79-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="3dd79-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="3dd79-124">displayLanguage :String</span></span>

<span data-ttu-id="3dd79-125">Obtém a localidade (idioma) no formato de marca de linguagem RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3dd79-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="3dd79-126">O valor `displayLanguage` reflete a configuração atual do **Idioma de Exibição** especificada em **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3dd79-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3dd79-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3dd79-127">Type:</span></span>

*   <span data-ttu-id="3dd79-128">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dd79-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dd79-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dd79-129">Requirements</span></span>

|<span data-ttu-id="3dd79-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dd79-130">Requirement</span></span>| <span data-ttu-id="3dd79-131">Valor</span><span class="sxs-lookup"><span data-stu-id="3dd79-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dd79-132">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dd79-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dd79-133">1.0</span><span class="sxs-lookup"><span data-stu-id="3dd79-133">1.0</span></span>|
|[<span data-ttu-id="3dd79-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dd79-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dd79-135">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3dd79-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dd79-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dd79-136">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="3dd79-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="3dd79-137">officeTheme :Object</span></span>

<span data-ttu-id="3dd79-138">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="3dd79-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="3dd79-139">Este membro não é suportado no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="3dd79-139">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3dd79-p102">Usando as cores de tema do Office, você pode coordenar o esquema de cores do seu suplemento com o tema atual do Office, selecionado pelo usuário em **Arquivo > Conta do Office > Tema da interface de usuário do Office**, que é aplicado a todos os aplicativos host do Office. Usar cores de tema do Office é apropriado para suplementos de painel de tarefas e email.</span><span class="sxs-lookup"><span data-stu-id="3dd79-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3dd79-142">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3dd79-142">Type:</span></span>

*   <span data-ttu-id="3dd79-143">Objeto</span><span class="sxs-lookup"><span data-stu-id="3dd79-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="3dd79-144">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3dd79-144">Properties:</span></span>

|<span data-ttu-id="3dd79-145">Nome</span><span class="sxs-lookup"><span data-stu-id="3dd79-145">Name</span></span>| <span data-ttu-id="3dd79-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dd79-146">Type</span></span>| <span data-ttu-id="3dd79-147">Descrição</span><span class="sxs-lookup"><span data-stu-id="3dd79-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="3dd79-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dd79-148">String</span></span>|<span data-ttu-id="3dd79-149">Obtém a cor do plano de fundo do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="3dd79-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="3dd79-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dd79-150">String</span></span>|<span data-ttu-id="3dd79-151">Obtém a cor de primeiro plano do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="3dd79-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="3dd79-152">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dd79-152">String</span></span>|<span data-ttu-id="3dd79-153">Obtém o tema do Office para controlar a cor do plano de fundo como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="3dd79-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="3dd79-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dd79-154">String</span></span>|<span data-ttu-id="3dd79-155">Obtém a cor de controle do corpo do tema do Office como um trio de cores hexadecimais.</span><span class="sxs-lookup"><span data-stu-id="3dd79-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3dd79-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dd79-156">Requirements</span></span>

|<span data-ttu-id="3dd79-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dd79-157">Requirement</span></span>| <span data-ttu-id="3dd79-158">Valor</span><span class="sxs-lookup"><span data-stu-id="3dd79-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dd79-159">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dd79-159">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dd79-160">1.3</span><span class="sxs-lookup"><span data-stu-id="3dd79-160">1.3</span></span>|
|[<span data-ttu-id="3dd79-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dd79-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dd79-162">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3dd79-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dd79-163">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dd79-163">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="3dd79-164">roamingSettings:[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="3dd79-164">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="3dd79-165">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email salvo na caixa de correio de um usuário.</span><span class="sxs-lookup"><span data-stu-id="3dd79-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3dd79-166">O objeto `RoamingSettings` permite armazenar e acessar dados para um suplemento de email armazenado na caixa de correio de um usuário, para que ele esteja disponível para esse complemento quando estiver sendo executado em qualquer aplicativo cliente host usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="3dd79-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3dd79-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3dd79-167">Type:</span></span>

*   [<span data-ttu-id="3dd79-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dd79-168">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3dd79-169">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dd79-169">Requirements</span></span>

|<span data-ttu-id="3dd79-170">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dd79-170">Requirement</span></span>| <span data-ttu-id="3dd79-171">Valor</span><span class="sxs-lookup"><span data-stu-id="3dd79-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dd79-172">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dd79-172">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3dd79-173">1.0</span><span class="sxs-lookup"><span data-stu-id="3dd79-173">1.0</span></span>|
|[<span data-ttu-id="3dd79-174">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3dd79-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3dd79-175">Restrito</span><span class="sxs-lookup"><span data-stu-id="3dd79-175">Restricted</span></span>|
|[<span data-ttu-id="3dd79-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dd79-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3dd79-177">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3dd79-177">Compose or read</span></span>|