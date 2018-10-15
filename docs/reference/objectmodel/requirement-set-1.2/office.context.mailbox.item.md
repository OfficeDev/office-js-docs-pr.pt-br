
# <a name="item"></a><span data-ttu-id="e5874-101">item</span><span class="sxs-lookup"><span data-stu-id="e5874-101">item</span></span>

### <span data-ttu-id="e5874-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e5874-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="e5874-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="e5874-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-106">Requirements</span></span>

|<span data-ttu-id="e5874-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-107">Requirement</span></span>| <span data-ttu-id="e5874-108">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-109">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-110">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-110">1.0</span></span>|
|[<span data-ttu-id="e5874-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="e5874-112">Restricted</span></span>|
|[<span data-ttu-id="e5874-113">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-114">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="e5874-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-115">Example</span></span>

<span data-ttu-id="e5874-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5874-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="e5874-117">Membros</span><span class="sxs-lookup"><span data-stu-id="e5874-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="e5874-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5874-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="e5874-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-121">Certos tipos de arquivos são bloqueados pelo Outlook devido a potenciais problemas de segurança e portanto não são retornados.</span><span class="sxs-lookup"><span data-stu-id="e5874-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e5874-122">Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="e5874-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-123">Type:</span></span>

*   <span data-ttu-id="e5874-124">Array. <[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5874-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-125">Requirements</span></span>

|<span data-ttu-id="e5874-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-126">Requirement</span></span>| <span data-ttu-id="e5874-127">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-128">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-129">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-129">1.0</span></span>|
|[<span data-ttu-id="e5874-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-131">ReadItem</span></span>|
|[<span data-ttu-id="e5874-132">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-133">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-134">Example</span></span>

<span data-ttu-id="e5874-135">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="e5874-136">cco:[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="e5874-137">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e5874-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="e5874-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-139">Type:</span></span>

*   [<span data-ttu-id="e5874-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="e5874-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e5874-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-141">Requirements</span></span>

|<span data-ttu-id="e5874-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-142">Requirement</span></span>| <span data-ttu-id="e5874-143">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-144">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-145">1.1</span><span class="sxs-lookup"><span data-stu-id="e5874-145">1.1</span></span>|
|[<span data-ttu-id="e5874-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-147">ReadItem</span></span>|
|[<span data-ttu-id="e5874-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-149">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="e5874-151">corpo:[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="e5874-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="e5874-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="e5874-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-153">Type:</span></span>

*   [<span data-ttu-id="e5874-154">Body</span><span class="sxs-lookup"><span data-stu-id="e5874-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="e5874-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-155">Requirements</span></span>

|<span data-ttu-id="e5874-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-156">Requirement</span></span>| <span data-ttu-id="e5874-157">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-158">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-159">1.1</span><span class="sxs-lookup"><span data-stu-id="e5874-159">1.1</span></span>|
|[<span data-ttu-id="e5874-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-161">ReadItem</span></span>|
|[<span data-ttu-id="e5874-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="e5874-164">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="e5874-165">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e5874-166">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-167">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-167">Read mode</span></span>

<span data-ttu-id="e5874-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="e5874-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-170">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-170">Compose mode</span></span>

<span data-ttu-id="e5874-171">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-172">Type:</span></span>

*   <span data-ttu-id="e5874-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-174">Requirements</span></span>

|<span data-ttu-id="e5874-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-175">Requirement</span></span>| <span data-ttu-id="e5874-176">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-177">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-178">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-178">1.0</span></span>|
|[<span data-ttu-id="e5874-179">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-180">ReadItem</span></span>|
|[<span data-ttu-id="e5874-181">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-182">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="e5874-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="e5874-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="e5874-185">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="e5874-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e5874-p108">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviá-la, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não será mais aplicável.</span><span class="sxs-lookup"><span data-stu-id="e5874-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e5874-p109">Para um novo item em um formulário de redação, o valor dessa propriedade é nulo. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="e5874-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-190">Type:</span></span>

*   <span data-ttu-id="e5874-191">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-192">Requirements</span></span>

|<span data-ttu-id="e5874-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-193">Requirement</span></span>| <span data-ttu-id="e5874-194">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-195">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-196">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-196">1.0</span></span>|
|[<span data-ttu-id="e5874-197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-198">ReadItem</span></span>|
|[<span data-ttu-id="e5874-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-200">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="e5874-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="e5874-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="e5874-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-204">Type:</span></span>

*   <span data-ttu-id="e5874-205">Data</span><span class="sxs-lookup"><span data-stu-id="e5874-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-206">Requirements</span></span>

|<span data-ttu-id="e5874-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-207">Requirement</span></span>| <span data-ttu-id="e5874-208">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-209">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-210">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-210">1.0</span></span>|
|[<span data-ttu-id="e5874-211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-212">ReadItem</span></span>|
|[<span data-ttu-id="e5874-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-214">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="e5874-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="e5874-216">dateTimeModified :Date</span></span>

<span data-ttu-id="e5874-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-219">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-220">Type:</span></span>

*   <span data-ttu-id="e5874-221">Data</span><span class="sxs-lookup"><span data-stu-id="e5874-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-222">Requirements</span></span>

|<span data-ttu-id="e5874-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-223">Requirement</span></span>| <span data-ttu-id="e5874-224">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-225">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-226">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-226">1.0</span></span>|
|[<span data-ttu-id="e5874-227">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-228">ReadItem</span></span>|
|[<span data-ttu-id="e5874-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-230">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-231">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="e5874-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5874-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="e5874-233">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="e5874-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e5874-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor da propriedade para a data e hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e5874-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-236">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-236">Read mode</span></span>

<span data-ttu-id="e5874-237">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="e5874-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-238">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-238">Compose mode</span></span>

<span data-ttu-id="e5874-239">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5874-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e5874-240">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC do servidor.</span><span class="sxs-lookup"><span data-stu-id="e5874-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-241">Type:</span></span>

*   <span data-ttu-id="e5874-242">Data | [Hora](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5874-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-243">Requirements</span></span>

|<span data-ttu-id="e5874-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-244">Requirement</span></span>| <span data-ttu-id="e5874-245">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-246">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-247">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-247">1.0</span></span>|
|[<span data-ttu-id="e5874-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-249">ReadItem</span></span>|
|[<span data-ttu-id="e5874-250">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-251">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-252">Example</span></span>

<span data-ttu-id="e5874-253">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5874-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="e5874-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5874-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5874-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="e5874-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="e5874-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-259">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e5874-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-260">Type:</span></span>

*   [<span data-ttu-id="e5874-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5874-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5874-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-262">Requirements</span></span>

|<span data-ttu-id="e5874-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-263">Requirement</span></span>| <span data-ttu-id="e5874-264">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-265">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-266">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-266">1.0</span></span>|
|[<span data-ttu-id="e5874-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-268">ReadItem</span></span>|
|[<span data-ttu-id="e5874-269">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-270">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="e5874-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="e5874-271">internetMessageId :String</span></span>

<span data-ttu-id="e5874-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-274">Type:</span></span>

*   <span data-ttu-id="e5874-275">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-276">Requirements</span></span>

|<span data-ttu-id="e5874-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-277">Requirement</span></span>| <span data-ttu-id="e5874-278">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-279">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-280">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-280">1.0</span></span>|
|[<span data-ttu-id="e5874-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-282">ReadItem</span></span>|
|[<span data-ttu-id="e5874-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-284">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="e5874-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="e5874-286">itemClass :String</span></span>

<span data-ttu-id="e5874-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e5874-p117">A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir estão as classes de mensagem padrão para itens de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="e5874-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-291">Type</span></span> | <span data-ttu-id="e5874-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-292">Description</span></span> | <span data-ttu-id="e5874-293">classe do item</span><span class="sxs-lookup"><span data-stu-id="e5874-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="e5874-294">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="e5874-294">Appointment items</span></span> | <span data-ttu-id="e5874-295">São itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="e5874-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="e5874-296">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="e5874-296">Message items</span></span> | <span data-ttu-id="e5874-297">Incluem mensagens de e-mail que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagens base.</span><span class="sxs-lookup"><span data-stu-id="e5874-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="e5874-298">Você pode criar classes de mensagens personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="e5874-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-299">Type:</span></span>

*   <span data-ttu-id="e5874-300">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-301">Requirements</span></span>

|<span data-ttu-id="e5874-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-302">Requirement</span></span>| <span data-ttu-id="e5874-303">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-304">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-305">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-305">1.0</span></span>|
|[<span data-ttu-id="e5874-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-307">ReadItem</span></span>|
|[<span data-ttu-id="e5874-308">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-309">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e5874-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="e5874-311">(nullable) itemId :String</span></span>

<span data-ttu-id="e5874-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-314">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="e5874-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e5874-315">A propriedade `itemId` não é idêntica à ID de entrada do Outlook ou à ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e5874-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e5874-316">Antes de fazer chamadas à API REST usando esse valor, ela deve ser convertida usando `Office.context.mailbox.convertToRestId`, que está disponível a partir do conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="e5874-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="e5874-317">Para obter mais detalhes, confira [Usar as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="e5874-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-318">Type:</span></span>

*   <span data-ttu-id="e5874-319">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-320">Requirements</span></span>

|<span data-ttu-id="e5874-321">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-321">Requirement</span></span>| <span data-ttu-id="e5874-322">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-323">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-324">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-324">1.0</span></span>|
|[<span data-ttu-id="e5874-325">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-326">ReadItem</span></span>|
|[<span data-ttu-id="e5874-327">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-328">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-329">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-329">Example</span></span>

<span data-ttu-id="e5874-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item a partir do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e5874-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="e5874-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e5874-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e5874-333">Obtém o tipo de item que uma instância representa.</span><span class="sxs-lookup"><span data-stu-id="e5874-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e5874-334">A propriedade `itemType` retorna um dos valores da enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-335">Type:</span></span>

*   [<span data-ttu-id="e5874-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e5874-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e5874-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-337">Requirements</span></span>

|<span data-ttu-id="e5874-338">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-338">Requirement</span></span>| <span data-ttu-id="e5874-339">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-340">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-341">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-341">1.0</span></span>|
|[<span data-ttu-id="e5874-342">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-343">ReadItem</span></span>|
|[<span data-ttu-id="e5874-344">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-345">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="e5874-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="e5874-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="e5874-348">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-349">Read mode</span></span>

<span data-ttu-id="e5874-350">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-351">Compose mode</span></span>

<span data-ttu-id="e5874-352">A propriedade `location` retorna um objeto `Location` que fornece métodos para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-353">Type:</span></span>

*   <span data-ttu-id="e5874-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="e5874-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-355">Requirements</span></span>

|<span data-ttu-id="e5874-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-356">Requirement</span></span>| <span data-ttu-id="e5874-357">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-358">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-359">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-359">1.0</span></span>|
|[<span data-ttu-id="e5874-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-361">ReadItem</span></span>|
|[<span data-ttu-id="e5874-362">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-363">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-364">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e5874-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="e5874-365">normalizedSubject :String</span></span>

<span data-ttu-id="e5874-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e5874-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="e5874-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-370">Type:</span></span>

*   <span data-ttu-id="e5874-371">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-372">Requirements</span></span>

|<span data-ttu-id="e5874-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-373">Requirement</span></span>| <span data-ttu-id="e5874-374">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-375">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-376">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-376">1.0</span></span>|
|[<span data-ttu-id="e5874-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-378">ReadItem</span></span>|
|[<span data-ttu-id="e5874-379">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-380">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-381">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="e5874-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="e5874-383">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="e5874-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e5874-384">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-385">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-385">Read mode</span></span>

<span data-ttu-id="e5874-386">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="e5874-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-387">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-387">Compose mode</span></span>

<span data-ttu-id="e5874-388">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="e5874-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-389">Type:</span></span>

*   <span data-ttu-id="e5874-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-391">Requirements</span></span>

|<span data-ttu-id="e5874-392">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-392">Requirement</span></span>| <span data-ttu-id="e5874-393">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-394">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-395">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-395">1.0</span></span>|
|[<span data-ttu-id="e5874-396">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-397">ReadItem</span></span>|
|[<span data-ttu-id="e5874-398">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-399">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-400">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="e5874-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5874-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5874-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-404">Type:</span></span>

*   [<span data-ttu-id="e5874-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5874-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5874-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-406">Requirements</span></span>

|<span data-ttu-id="e5874-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-407">Requirement</span></span>| <span data-ttu-id="e5874-408">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-409">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-410">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-410">1.0</span></span>|
|[<span data-ttu-id="e5874-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-412">ReadItem</span></span>|
|[<span data-ttu-id="e5874-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-414">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="e5874-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="e5874-417">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="e5874-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e5874-418">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-419">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-419">Read mode</span></span>

<span data-ttu-id="e5874-420">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="e5874-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-421">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-421">Compose mode</span></span>

<span data-ttu-id="e5874-422">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="e5874-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-423">Type:</span></span>

*   <span data-ttu-id="e5874-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-425">Requirements</span></span>

|<span data-ttu-id="e5874-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-426">Requirement</span></span>| <span data-ttu-id="e5874-427">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-428">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-429">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-429">1.0</span></span>|
|[<span data-ttu-id="e5874-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-431">ReadItem</span></span>|
|[<span data-ttu-id="e5874-432">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-433">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="e5874-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5874-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5874-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e5874-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e5874-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegador, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="e5874-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-440">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="e5874-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-441">Type:</span></span>

*   [<span data-ttu-id="e5874-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5874-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5874-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-443">Requirements</span></span>

|<span data-ttu-id="e5874-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-444">Requirement</span></span>| <span data-ttu-id="e5874-445">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-446">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-447">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-447">1.0</span></span>|
|[<span data-ttu-id="e5874-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-449">ReadItem</span></span>|
|[<span data-ttu-id="e5874-450">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-451">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="e5874-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5874-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="e5874-454">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="e5874-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e5874-p128">A propriedade `start` é expressa como um valor de data e valor temporal no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e5874-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-457">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-457">Read mode</span></span>

<span data-ttu-id="e5874-458">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="e5874-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-459">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-459">Compose mode</span></span>

<span data-ttu-id="e5874-460">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5874-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e5874-461">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="e5874-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-462">Type:</span></span>

*   <span data-ttu-id="e5874-463">Data | [Hora](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5874-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-464">Requirements</span></span>

|<span data-ttu-id="e5874-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-465">Requirement</span></span>| <span data-ttu-id="e5874-466">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-467">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-468">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-468">1.0</span></span>|
|[<span data-ttu-id="e5874-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-470">ReadItem</span></span>|
|[<span data-ttu-id="e5874-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-472">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-473">Example</span></span>

<span data-ttu-id="e5874-474">O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="e5874-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="e5874-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e5874-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="e5874-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="e5874-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e5874-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de e-mail.</span><span class="sxs-lookup"><span data-stu-id="e5874-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-478">Read mode</span></span>

<span data-ttu-id="e5874-p129">A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="e5874-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="e5874-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-481">Compose mode</span></span>

<span data-ttu-id="e5874-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="e5874-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5874-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-483">Type:</span></span>

*   <span data-ttu-id="e5874-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e5874-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-485">Requirements</span></span>

|<span data-ttu-id="e5874-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-486">Requirement</span></span>| <span data-ttu-id="e5874-487">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-488">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-489">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-489">1.0</span></span>|
|[<span data-ttu-id="e5874-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-491">ReadItem</span></span>|
|[<span data-ttu-id="e5874-492">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-493">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="e5874-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="e5874-495">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e5874-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5874-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-497">Read mode</span></span>

<span data-ttu-id="e5874-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **To** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="e5874-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5874-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="e5874-500">Compose mode</span></span>

<span data-ttu-id="e5874-501">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **To** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e5874-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="e5874-502">Type:</span></span>

*   <span data-ttu-id="e5874-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5874-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-504">Requirements</span></span>

|<span data-ttu-id="e5874-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-505">Requirement</span></span>| <span data-ttu-id="e5874-506">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-507">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-508">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-508">1.0</span></span>|
|[<span data-ttu-id="e5874-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-510">ReadItem</span></span>|
|[<span data-ttu-id="e5874-511">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-512">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="e5874-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="e5874-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e5874-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5874-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5874-516">Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.</span><span class="sxs-lookup"><span data-stu-id="e5874-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e5874-517">O método `addFileAttachmentAsync` carrega o arquivo da URI especificada e o anexa ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="e5874-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e5874-518">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e5874-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-519">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-519">Parameters:</span></span>

|<span data-ttu-id="e5874-520">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-520">Name</span></span>| <span data-ttu-id="e5874-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-521">Type</span></span>| <span data-ttu-id="e5874-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-522">Attributes</span></span>| <span data-ttu-id="e5874-523">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="e5874-524">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-524">String</span></span>||<span data-ttu-id="e5874-p132">O URI que fornece a localização do arquivo anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e5874-527">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-527">String</span></span>||<span data-ttu-id="e5874-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e5874-530">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-530">Object</span></span>| <span data-ttu-id="e5874-531">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-531">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-532">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5874-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-533">Object</span></span>| <span data-ttu-id="e5874-534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-534">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-535">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5874-536">function</span><span class="sxs-lookup"><span data-stu-id="e5874-536">function</span></span>| <span data-ttu-id="e5874-537">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-537">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-538">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5874-539">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5874-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5874-540">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="e5874-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5874-541">Erros</span><span class="sxs-lookup"><span data-stu-id="e5874-541">Errors</span></span>

| <span data-ttu-id="e5874-542">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e5874-542">Error code</span></span> | <span data-ttu-id="e5874-543">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="e5874-544">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="e5874-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="e5874-545">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="e5874-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e5874-546">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="e5874-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-547">Requirements</span></span>

|<span data-ttu-id="e5874-548">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-548">Requirement</span></span>| <span data-ttu-id="e5874-549">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-550">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-551">1.1</span><span class="sxs-lookup"><span data-stu-id="e5874-551">1.1</span></span>|
|[<span data-ttu-id="e5874-552">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5874-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5874-554">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-555">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-556">Example</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e5874-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5874-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5874-558">Adiciona um item do Exchange, como uma mensagem, como um anexo à mensagem ou ao compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e5874-p134">O método `addItemAttachmentAsync` anexa o item com o identificador especificado do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="e5874-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e5874-562">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="e5874-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e5874-563">Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a outros itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="e5874-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-564">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-564">Parameters:</span></span>

|<span data-ttu-id="e5874-565">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-565">Name</span></span>| <span data-ttu-id="e5874-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-566">Type</span></span>| <span data-ttu-id="e5874-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-567">Attributes</span></span>| <span data-ttu-id="e5874-568">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="e5874-569">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-569">String</span></span>||<span data-ttu-id="e5874-p135">O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e5874-572">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-572">String</span></span>||<span data-ttu-id="e5874-p136">O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e5874-575">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-575">Object</span></span>| <span data-ttu-id="e5874-576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-576">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-577">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5874-578">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-578">Object</span></span>| <span data-ttu-id="e5874-579">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-579">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-580">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5874-581">function</span><span class="sxs-lookup"><span data-stu-id="e5874-581">function</span></span>| <span data-ttu-id="e5874-582">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-582">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-583">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5874-584">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5874-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5874-585">Se não for possível adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` com a descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="e5874-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5874-586">Erros</span><span class="sxs-lookup"><span data-stu-id="e5874-586">Errors</span></span>

| <span data-ttu-id="e5874-587">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e5874-587">Error code</span></span> | <span data-ttu-id="e5874-588">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e5874-589">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="e5874-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-590">Requirements</span></span>

|<span data-ttu-id="e5874-591">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-591">Requirement</span></span>| <span data-ttu-id="e5874-592">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-593">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-594">1.1</span><span class="sxs-lookup"><span data-stu-id="e5874-594">1.1</span></span>|
|[<span data-ttu-id="e5874-595">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5874-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5874-597">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-598">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-599">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-599">Example</span></span>

<span data-ttu-id="e5874-600">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="e5874-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="e5874-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e5874-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="e5874-602">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="e5874-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-603">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5874-604">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="e5874-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5874-605">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e5874-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e5874-p137">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="e5874-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-609">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-609">Parameters:</span></span>

|<span data-ttu-id="e5874-610">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-610">Name</span></span>| <span data-ttu-id="e5874-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-611">Type</span></span>| <span data-ttu-id="e5874-612">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e5874-613">String | Object</span><span class="sxs-lookup"><span data-stu-id="e5874-613">String &#124; Object</span></span>| |<span data-ttu-id="e5874-p138">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e5874-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5874-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="e5874-616">**OR**</span></span><br/><span data-ttu-id="e5874-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="e5874-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e5874-619">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-619">String</span></span> | <span data-ttu-id="e5874-620">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-620">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-p140">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e5874-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="e5874-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e5874-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-624">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-625">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="e5874-626">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-626">String</span></span> | | <span data-ttu-id="e5874-p141">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="e5874-629">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-629">String</span></span> | | <span data-ttu-id="e5874-630">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="e5874-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="e5874-631">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-631">String</span></span> | | <span data-ttu-id="e5874-p142">Usado somente se `type` estiver definido como `file`. O URI da localização para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e5874-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="e5874-634">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-634">String</span></span> | | <span data-ttu-id="e5874-p143">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="e5874-638">função</span><span class="sxs-lookup"><span data-stu-id="e5874-638">function</span></span> | <span data-ttu-id="e5874-639">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-639">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-640">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-641">Requirements</span></span>

|<span data-ttu-id="e5874-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-642">Requirement</span></span>| <span data-ttu-id="e5874-643">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-644">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-645">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-645">1.0</span></span>|
|[<span data-ttu-id="e5874-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-647">ReadItem</span></span>|
|[<span data-ttu-id="e5874-648">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-649">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5874-650">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e5874-650">Examples</span></span>

<span data-ttu-id="e5874-651">O código a seguir passa uma sequência de caracteres para a função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="e5874-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e5874-652">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="e5874-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e5874-653">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="e5874-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5874-654">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e5874-654">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="e5874-655">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-655">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="e5874-656">Resposta com o corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="e5874-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e5874-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="e5874-658">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="e5874-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-659">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5874-660">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="e5874-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5874-661">Se qualquer um dos parâmetros do tipo sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e5874-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e5874-p144">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="e5874-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-665">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-665">Parameters:</span></span>

|<span data-ttu-id="e5874-666">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-666">Name</span></span>| <span data-ttu-id="e5874-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-667">Type</span></span>| <span data-ttu-id="e5874-668">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e5874-669">String | Object</span><span class="sxs-lookup"><span data-stu-id="e5874-669">String &#124; Object</span></span>| | <span data-ttu-id="e5874-p145">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e5874-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5874-672">**OU**</span><span class="sxs-lookup"><span data-stu-id="e5874-672">**OR**</span></span><br/><span data-ttu-id="e5874-p146">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="e5874-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e5874-675">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-675">String</span></span> | <span data-ttu-id="e5874-676">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-676">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-p147">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e5874-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="e5874-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e5874-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-680">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-681">Uma matriz de objetos JSON que são anexos de arquivo ou de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="e5874-682">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-682">String</span></span> | | <span data-ttu-id="e5874-p148">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="e5874-685">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-685">String</span></span> | | <span data-ttu-id="e5874-686">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="e5874-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="e5874-687">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-687">String</span></span> | | <span data-ttu-id="e5874-p149">Usado somente se `type` estiver definido como `file`. O URI da localização para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e5874-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="e5874-690">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-690">String</span></span> | | <span data-ttu-id="e5874-p150">Usado somente se `type` estiver definido como `item`. O ID do item do anexo no EWS. É uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="e5874-694">função</span><span class="sxs-lookup"><span data-stu-id="e5874-694">function</span></span> | <span data-ttu-id="e5874-695">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-695">&lt;optional&gt;</span></span> | <span data-ttu-id="e5874-696">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-697">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-697">Requirements</span></span>

|<span data-ttu-id="e5874-698">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-698">Requirement</span></span>| <span data-ttu-id="e5874-699">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-700">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-700">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-701">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-701">1.0</span></span>|
|[<span data-ttu-id="e5874-702">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-703">ReadItem</span></span>|
|[<span data-ttu-id="e5874-704">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-705">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5874-706">Exemplos</span><span class="sxs-lookup"><span data-stu-id="e5874-706">Examples</span></span>

<span data-ttu-id="e5874-707">O código a seguir passa uma sequência de caracteres para a função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="e5874-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e5874-708">Resposta com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="e5874-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e5874-709">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="e5874-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5874-710">Resposta com o corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e5874-710">Reply with a body and a file attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="e5874-711">Resposta com o corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e5874-711">Reply with a body and an item attachment.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="e5874-712">Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="e5874-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e5874-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="e5874-714">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e5874-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-715">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-716">Requirements</span></span>

|<span data-ttu-id="e5874-717">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-717">Requirement</span></span>| <span data-ttu-id="e5874-718">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-719">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-720">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-720">1.0</span></span>|
|[<span data-ttu-id="e5874-721">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-722">ReadItem</span></span>|
|[<span data-ttu-id="e5874-723">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-724">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-725">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-725">Returns:</span></span>

<span data-ttu-id="e5874-726">Tipo: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e5874-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e5874-727">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-727">Example</span></span>

<span data-ttu-id="e5874-728">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-728">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="e5874-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5874-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5874-730">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e5874-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-731">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-732">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-732">Parameters:</span></span>

|<span data-ttu-id="e5874-733">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-733">Name</span></span>| <span data-ttu-id="e5874-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-734">Type</span></span>| <span data-ttu-id="e5874-735">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="e5874-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e5874-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="e5874-737">Um dos valores da enumeração EntityType.</span><span class="sxs-lookup"><span data-stu-id="e5874-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5874-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-738">Requirements</span></span>

|<span data-ttu-id="e5874-739">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-739">Requirement</span></span>| <span data-ttu-id="e5874-740">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-741">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-742">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-742">1.0</span></span>|
|[<span data-ttu-id="e5874-743">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-744">Restrito</span><span class="sxs-lookup"><span data-stu-id="e5874-744">Restricted</span></span>|
|[<span data-ttu-id="e5874-745">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-746">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-747">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-747">Returns:</span></span>

<span data-ttu-id="e5874-748">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="e5874-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e5874-749">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="e5874-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="e5874-750">Caso contrário, o tipo dos objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="e5874-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e5874-751">Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="e5874-752">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="e5874-752">Value of `entityType`</span></span> | <span data-ttu-id="e5874-753">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="e5874-753">Type of objects in returned array</span></span> | <span data-ttu-id="e5874-754">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="e5874-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="e5874-755">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-755">String</span></span> | <span data-ttu-id="e5874-756">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e5874-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="e5874-757">Contact</span><span class="sxs-lookup"><span data-stu-id="e5874-757">Contact</span></span> | <span data-ttu-id="e5874-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5874-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="e5874-759">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-759">String</span></span> | <span data-ttu-id="e5874-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5874-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="e5874-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5874-761">MeetingSuggestion</span></span> | <span data-ttu-id="e5874-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5874-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="e5874-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e5874-763">PhoneNumber</span></span> | <span data-ttu-id="e5874-764">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e5874-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="e5874-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5874-765">TaskSuggestion</span></span> | <span data-ttu-id="e5874-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5874-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="e5874-767">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-767">String</span></span> | <span data-ttu-id="e5874-768">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="e5874-768">**Restricted**</span></span> |

<span data-ttu-id="e5874-769">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5874-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="e5874-770">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-770">Example</span></span>

<span data-ttu-id="e5874-771">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa os endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="e5874-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```JavaScript
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="e5874-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5874-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5874-773">Retorna entidades conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e5874-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-774">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5874-775">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor especificado no elemento `FilterName` .</span><span class="sxs-lookup"><span data-stu-id="e5874-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-776">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-776">Parameters:</span></span>

|<span data-ttu-id="e5874-777">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-777">Name</span></span>| <span data-ttu-id="e5874-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-778">Type</span></span>| <span data-ttu-id="e5874-779">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e5874-780">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-780">String</span></span>|<span data-ttu-id="e5874-781">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="e5874-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5874-782">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-782">Requirements</span></span>

|<span data-ttu-id="e5874-783">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-783">Requirement</span></span>| <span data-ttu-id="e5874-784">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-785">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-785">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-786">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-786">1.0</span></span>|
|[<span data-ttu-id="e5874-787">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-788">ReadItem</span></span>|
|[<span data-ttu-id="e5874-789">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-790">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-791">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-791">Returns:</span></span>

<span data-ttu-id="e5874-p152">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="e5874-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e5874-794">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5874-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="e5874-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e5874-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e5874-796">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e5874-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-797">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5874-p153">O método `getRegExMatches` retorna as sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e5874-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e5874-801">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="e5874-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e5874-802">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="e5874-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="e5874-p154">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade body de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="e5874-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5874-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-805">Requirements</span></span>

|<span data-ttu-id="e5874-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-806">Requirement</span></span>| <span data-ttu-id="e5874-807">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-808">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-809">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-809">1.0</span></span>|
|[<span data-ttu-id="e5874-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-811">ReadItem</span></span>|
|[<span data-ttu-id="e5874-812">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-813">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-814">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-814">Returns:</span></span>

<span data-ttu-id="e5874-p155">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="e5874-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e5874-817">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5874-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5874-818">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5874-819">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-819">Example</span></span>

<span data-ttu-id="e5874-820">O exemplo a seguir mostra como acessar a matriz de correspondências para os <rule>elementos `fruits` e `veggies` da expressão regular que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="e5874-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e5874-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e5874-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e5874-822">Retorna valores do tipo sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e5874-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5874-823">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e5874-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5874-824">O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="e5874-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e5874-p156">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="e5874-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-827">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-827">Parameters:</span></span>

|<span data-ttu-id="e5874-828">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-828">Name</span></span>| <span data-ttu-id="e5874-829">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-829">Type</span></span>| <span data-ttu-id="e5874-830">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e5874-831">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-831">String</span></span>|<span data-ttu-id="e5874-832">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="e5874-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5874-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-833">Requirements</span></span>

|<span data-ttu-id="e5874-834">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-834">Requirement</span></span>| <span data-ttu-id="e5874-835">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-836">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-837">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-837">1.0</span></span>|
|[<span data-ttu-id="e5874-838">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-839">ReadItem</span></span>|
|[<span data-ttu-id="e5874-840">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-841">Leitura</span><span class="sxs-lookup"><span data-stu-id="e5874-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-842">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-842">Returns:</span></span>

<span data-ttu-id="e5874-843">Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="e5874-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="e5874-844">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5874-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5874-845">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e5874-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5874-846">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e5874-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e5874-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e5874-848">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e5874-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e5874-p157">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="e5874-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-851">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-851">Parameters:</span></span>

|<span data-ttu-id="e5874-852">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-852">Name</span></span>| <span data-ttu-id="e5874-853">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-853">Type</span></span>| <span data-ttu-id="e5874-854">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-854">Attributes</span></span>| <span data-ttu-id="e5874-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="e5874-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e5874-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e5874-p158">Solicita um formato para os dados. Se for Text, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="e5874-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="e5874-860">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-860">Object</span></span>| <span data-ttu-id="e5874-861">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-861">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-862">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5874-863">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-863">Object</span></span>| <span data-ttu-id="e5874-864">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-864">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-865">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5874-866">function</span><span class="sxs-lookup"><span data-stu-id="e5874-866">function</span></span>||<span data-ttu-id="e5874-867">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5874-868">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="e5874-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e5874-869">Para acessar a propriedade de origem de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="e5874-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5874-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-870">Requirements</span></span>

|<span data-ttu-id="e5874-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-871">Requirement</span></span>| <span data-ttu-id="e5874-872">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-873">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-874">1.2</span><span class="sxs-lookup"><span data-stu-id="e5874-874">1.2</span></span>|
|[<span data-ttu-id="e5874-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5874-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5874-877">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-878">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5874-879">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e5874-879">Returns:</span></span>

<span data-ttu-id="e5874-880">Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="e5874-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="e5874-881">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5874-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5874-882">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5874-883">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-883">Example</span></span>

```JavaScript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e5874-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e5874-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e5874-885">Carrega de forma assíncrona as propriedades personalizadas desse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="e5874-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e5874-p160">As propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item e o suplemento atuais. As propriedades personalizadas não são criptografadas no item, portanto, isto não deve ser usado como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="e5874-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-889">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-889">Parameters:</span></span>

|<span data-ttu-id="e5874-890">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-890">Name</span></span>| <span data-ttu-id="e5874-891">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-891">Type</span></span>| <span data-ttu-id="e5874-892">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-892">Attributes</span></span>| <span data-ttu-id="e5874-893">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e5874-894">function</span><span class="sxs-lookup"><span data-stu-id="e5874-894">function</span></span>||<span data-ttu-id="e5874-895">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5874-896">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e5874-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e5874-897">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas no servidor.</span><span class="sxs-lookup"><span data-stu-id="e5874-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="e5874-898">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-898">Object</span></span>| <span data-ttu-id="e5874-899">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-899">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-900">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="e5874-901">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5874-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-902">Requirements</span></span>

|<span data-ttu-id="e5874-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-903">Requirement</span></span>| <span data-ttu-id="e5874-904">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-905">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-906">1.0</span><span class="sxs-lookup"><span data-stu-id="e5874-906">1.0</span></span>|
|[<span data-ttu-id="e5874-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5874-908">ReadItem</span></span>|
|[<span data-ttu-id="e5874-909">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-910">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="e5874-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-911">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-911">Example</span></span>

<span data-ttu-id="e5874-p163">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e5874-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e5874-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5874-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e5874-916">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e5874-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e5874-p164">O método `removeAttachmentAsync` remove do item o anexo com o identificador especificado. Conforme as práticas recomendadas, você deve usar o identificador do anexo para remover o anexo apenas se o mesmo aplicativo de email tiver inserido o anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador de anexos é válido somente dentro da mesma sessão. Uma sessão é considerada encerrada quando o usuário fecha o aplicativo, ou se o usuário começa a escrever um email em um formulário embutido e, em seguida, abre o mesmo formulário em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="e5874-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-921">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-921">Parameters:</span></span>

|<span data-ttu-id="e5874-922">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-922">Name</span></span>| <span data-ttu-id="e5874-923">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-923">Type</span></span>| <span data-ttu-id="e5874-924">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-924">Attributes</span></span>| <span data-ttu-id="e5874-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="e5874-926">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-926">String</span></span>||<span data-ttu-id="e5874-p165">O identificador do anexo a ser removido. O comprimento máximo da sequência de caracteres é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e5874-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="e5874-929">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-929">Object</span></span>| <span data-ttu-id="e5874-930">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-930">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-931">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5874-932">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-932">Object</span></span>| <span data-ttu-id="e5874-933">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-933">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-934">Os desenvolvedores podem fornecer qualquer objeto que desejarem para acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5874-935">function</span><span class="sxs-lookup"><span data-stu-id="e5874-935">function</span></span>| <span data-ttu-id="e5874-936">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-936">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-937">Quando o método for concluído, a função passada no parâmetro `callback` será chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5874-938">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="e5874-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5874-939">Erros</span><span class="sxs-lookup"><span data-stu-id="e5874-939">Errors</span></span>

| <span data-ttu-id="e5874-940">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e5874-940">Error code</span></span> | <span data-ttu-id="e5874-941">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="e5874-942">O identificador do anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="e5874-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-943">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-943">Requirements</span></span>

|<span data-ttu-id="e5874-944">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-944">Requirement</span></span>| <span data-ttu-id="e5874-945">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-946">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-947">1.1</span><span class="sxs-lookup"><span data-stu-id="e5874-947">1.1</span></span>|
|[<span data-ttu-id="e5874-948">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5874-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5874-950">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-951">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-952">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-952">Example</span></span>

<span data-ttu-id="e5874-953">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="e5874-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e5874-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e5874-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e5874-955">Insere dados no corpo ou no assunto de uma mensagem de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="e5874-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e5874-p166">O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou no assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou do assunto, um erro será retornado. Após a inserção, o cursor será posicionado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="e5874-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5874-959">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="e5874-959">Parameters:</span></span>

|<span data-ttu-id="e5874-960">Nome</span><span class="sxs-lookup"><span data-stu-id="e5874-960">Name</span></span>| <span data-ttu-id="e5874-961">Tipo</span><span class="sxs-lookup"><span data-stu-id="e5874-961">Type</span></span>| <span data-ttu-id="e5874-962">Atributos</span><span class="sxs-lookup"><span data-stu-id="e5874-962">Attributes</span></span>| <span data-ttu-id="e5874-963">Descrição</span><span class="sxs-lookup"><span data-stu-id="e5874-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e5874-964">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="e5874-964">String</span></span>||<span data-ttu-id="e5874-p167">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="e5874-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="e5874-968">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-968">Object</span></span>| <span data-ttu-id="e5874-969">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-969">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-970">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e5874-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5874-971">Objeto</span><span class="sxs-lookup"><span data-stu-id="e5874-971">Object</span></span>| <span data-ttu-id="e5874-972">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-972">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-973">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e5874-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="e5874-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e5874-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="e5874-975">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5874-975">&lt;optional&gt;</span></span>|<span data-ttu-id="e5874-p168">Se for `text` , o estilo atual será aplicado no Outlook Web App e no Outlook. Se o campo for um editor HTML, somente os dados de texto serão inseridos, mesmo que os dados estejam em HTML.</span><span class="sxs-lookup"><span data-stu-id="e5874-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e5874-p169">Se for `html` e o campo for compatível com HTML (e o assunto não), o estilo atual será aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, um erro `InvalidDataFormat` será retornado.</span><span class="sxs-lookup"><span data-stu-id="e5874-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e5874-980">Se `coercionType` não estiver definido, o resultado dependerá do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="e5874-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="e5874-981">função</span><span class="sxs-lookup"><span data-stu-id="e5874-981">function</span></span>||<span data-ttu-id="e5874-982">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e5874-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5874-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e5874-983">Requirements</span></span>

|<span data-ttu-id="e5874-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="e5874-984">Requirement</span></span>| <span data-ttu-id="e5874-985">Valor</span><span class="sxs-lookup"><span data-stu-id="e5874-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5874-986">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e5874-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5874-987">1.2</span><span class="sxs-lookup"><span data-stu-id="e5874-987">1.2</span></span>|
|[<span data-ttu-id="e5874-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e5874-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5874-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5874-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5874-990">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e5874-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5874-991">Redigir</span><span class="sxs-lookup"><span data-stu-id="e5874-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5874-992">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e5874-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```