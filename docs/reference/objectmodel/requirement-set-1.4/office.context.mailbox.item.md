
# <a name="item"></a><span data-ttu-id="fe079-101">item</span><span class="sxs-lookup"><span data-stu-id="fe079-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="fe079-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fe079-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="fe079-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="fe079-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-105">Requirements</span></span>

|<span data-ttu-id="fe079-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-106">Requirement</span></span>| <span data-ttu-id="fe079-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-109">1.0</span></span>|
|[<span data-ttu-id="fe079-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="fe079-111">Restricted</span></span>|
|[<span data-ttu-id="fe079-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-113">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="fe079-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-114">Example</span></span>

<span data-ttu-id="fe079-115">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe079-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
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

### <a name="members"></a><span data-ttu-id="fe079-116">Membros</span><span class="sxs-lookup"><span data-stu-id="fe079-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="fe079-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fe079-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="fe079-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-120">Certos tipos de arquivos são bloqueados pelo Outlook devido a possíveis problemas de segurança e, portanto, não retornam.</span><span class="sxs-lookup"><span data-stu-id="fe079-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fe079-121">Para saber mais, consulte [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="fe079-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-122">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-122">Type:</span></span>

*   <span data-ttu-id="fe079-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fe079-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-124">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-124">Requirements</span></span>

|<span data-ttu-id="fe079-125">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-125">Requirement</span></span>| <span data-ttu-id="fe079-126">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-127">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-128">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-128">1.0</span></span>|
|[<span data-ttu-id="fe079-129">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-130">ReadItem</span></span>|
|[<span data-ttu-id="fe079-131">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-133">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-133">Example</span></span>

<span data-ttu-id="fe079-134">O código a seguir cria uma cadeia de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="fe079-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="fe079-136">Obtém um objeto que fornece métodos para acessar ou atualizar a linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fe079-137">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fe079-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-138">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-138">Type:</span></span>

*   [<span data-ttu-id="fe079-139">Destinatários</span><span class="sxs-lookup"><span data-stu-id="fe079-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fe079-140">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-140">Requirements</span></span>

|<span data-ttu-id="fe079-141">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-141">Requirement</span></span>| <span data-ttu-id="fe079-142">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-143">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-144">1.1</span><span class="sxs-lookup"><span data-stu-id="fe079-144">1.1</span></span>|
|[<span data-ttu-id="fe079-145">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-146">ReadItem</span></span>|
|[<span data-ttu-id="fe079-147">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-148">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-149">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-149">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="fe079-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="fe079-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="fe079-151">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="fe079-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-152">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-152">Type:</span></span>

*   [<span data-ttu-id="fe079-153">Corpo</span><span class="sxs-lookup"><span data-stu-id="fe079-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="fe079-154">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-154">Requirements</span></span>

|<span data-ttu-id="fe079-155">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-155">Requirement</span></span>| <span data-ttu-id="fe079-156">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-157">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-158">1.1</span><span class="sxs-lookup"><span data-stu-id="fe079-158">1.1</span></span>|
|[<span data-ttu-id="fe079-159">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-160">ReadItem</span></span>|
|[<span data-ttu-id="fe079-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-162">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="fe079-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="fe079-164">Fornece acesso aos destinatários na linha Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fe079-165">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-166">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-166">Read mode</span></span>

<span data-ttu-id="fe079-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="fe079-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-169">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-169">Compose mode</span></span>

<span data-ttu-id="fe079-170">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-171">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-171">Type:</span></span>

*   <span data-ttu-id="fe079-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-173">Requirements</span></span>

|<span data-ttu-id="fe079-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-174">Requirement</span></span>| <span data-ttu-id="fe079-175">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-177">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-177">1.0</span></span>|
|[<span data-ttu-id="fe079-178">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-179">ReadItem</span></span>|
|[<span data-ttu-id="fe079-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-181">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-182">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-182">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fe079-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fe079-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="fe079-184">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="fe079-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fe079-p107">Você pode obter um número inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas em formulários de composição. Se, posteriormente, o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="fe079-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fe079-p108">Você obtém nulo para esta propriedade para um novo item em um formulário de composição. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="fe079-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-189">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-189">Type:</span></span>

*   <span data-ttu-id="fe079-190">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fe079-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-191">Requirements</span></span>

|<span data-ttu-id="fe079-192">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-192">Requirement</span></span>| <span data-ttu-id="fe079-193">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-195">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-195">1.0</span></span>|
|[<span data-ttu-id="fe079-196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-197">ReadItem</span></span>|
|[<span data-ttu-id="fe079-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-199">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="fe079-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fe079-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="fe079-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-203">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-203">Type:</span></span>

*   <span data-ttu-id="fe079-204">Data</span><span class="sxs-lookup"><span data-stu-id="fe079-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-205">Requirements</span></span>

|<span data-ttu-id="fe079-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-206">Requirement</span></span>| <span data-ttu-id="fe079-207">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-209">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-209">1.0</span></span>|
|[<span data-ttu-id="fe079-210">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-211">ReadItem</span></span>|
|[<span data-ttu-id="fe079-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-213">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-214">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fe079-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fe079-215">dateTimeModified :Date</span></span>

<span data-ttu-id="fe079-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-218">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-219">Type:</span></span>

*   <span data-ttu-id="fe079-220">Data</span><span class="sxs-lookup"><span data-stu-id="fe079-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-221">Requirements</span></span>

|<span data-ttu-id="fe079-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-222">Requirement</span></span>| <span data-ttu-id="fe079-223">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-225">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-225">1.0</span></span>|
|[<span data-ttu-id="fe079-226">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-227">ReadItem</span></span>|
|[<span data-ttu-id="fe079-228">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-229">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-230">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-230">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="fe079-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe079-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="fe079-232">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="fe079-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fe079-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="fe079-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-235">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-235">Read mode</span></span>

<span data-ttu-id="fe079-236">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="fe079-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-237">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-237">Compose mode</span></span>

<span data-ttu-id="fe079-238">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe079-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fe079-239">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="fe079-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-240">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-240">Type:</span></span>

*   <span data-ttu-id="fe079-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe079-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-242">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-242">Requirements</span></span>

|<span data-ttu-id="fe079-243">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-243">Requirement</span></span>| <span data-ttu-id="fe079-244">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-245">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-246">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-246">1.0</span></span>|
|[<span data-ttu-id="fe079-247">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-248">ReadItem</span></span>|
|[<span data-ttu-id="fe079-249">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-250">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-251">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-251">Example</span></span>

<span data-ttu-id="fe079-252">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe079-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="fe079-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe079-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe079-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="fe079-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="fe079-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-258">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fe079-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-259">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-259">Type:</span></span>

*   [<span data-ttu-id="fe079-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe079-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe079-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-261">Requirements</span></span>

|<span data-ttu-id="fe079-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-262">Requirement</span></span>| <span data-ttu-id="fe079-263">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-264">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-265">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-265">1.0</span></span>|
|[<span data-ttu-id="fe079-266">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-267">ReadItem</span></span>|
|[<span data-ttu-id="fe079-268">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-269">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="fe079-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fe079-270">internetMessageId :String</span></span>

<span data-ttu-id="fe079-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-273">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-273">Type:</span></span>

*   <span data-ttu-id="fe079-274">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fe079-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-275">Requirements</span></span>

|<span data-ttu-id="fe079-276">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-276">Requirement</span></span>| <span data-ttu-id="fe079-277">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-279">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-279">1.0</span></span>|
|[<span data-ttu-id="fe079-280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-281">ReadItem</span></span>|
|[<span data-ttu-id="fe079-282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-283">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-284">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-284">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fe079-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fe079-285">itemClass :String</span></span>

<span data-ttu-id="fe079-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fe079-p116">A propriedade `itemClass` especifica a classe da mensagem do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="fe079-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-290">Type</span></span> | <span data-ttu-id="fe079-291">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-291">Description</span></span> | <span data-ttu-id="fe079-292">classe de item</span><span class="sxs-lookup"><span data-stu-id="fe079-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="fe079-293">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="fe079-293">Appointment items</span></span> | <span data-ttu-id="fe079-294">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="fe079-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="fe079-295">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="fe079-295">Message items</span></span> | <span data-ttu-id="fe079-296">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos, que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="fe079-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="fe079-297">Você pode criar classes de mensagem personalizadas que estendem uma classe de mensagem padrão, por exemplo, uma classe de mensagem de compromisso `IPM.Appointment.Contoso` personalizada.</span><span class="sxs-lookup"><span data-stu-id="fe079-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-298">Type:</span></span>

*   <span data-ttu-id="fe079-299">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fe079-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-300">Requirements</span></span>

|<span data-ttu-id="fe079-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-301">Requirement</span></span>| <span data-ttu-id="fe079-302">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-303">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-304">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-304">1.0</span></span>|
|[<span data-ttu-id="fe079-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-306">ReadItem</span></span>|
|[<span data-ttu-id="fe079-307">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-308">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-309">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fe079-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fe079-310">(nullable) itemId :String</span></span>

<span data-ttu-id="fe079-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-313">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="fe079-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fe079-314">A propriedade `itemId` não é idêntica à ID de Entrada do Outlook ou a ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fe079-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fe079-315">Antes de fazer chamadas API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="fe079-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fe079-316">Para saber mais, consulte [Usar APIs REST do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="fe079-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="fe079-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório o identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no servidor, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-319">Type:</span></span>

*   <span data-ttu-id="fe079-320">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fe079-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-321">Requirements</span></span>

|<span data-ttu-id="fe079-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-322">Requirement</span></span>| <span data-ttu-id="fe079-323">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-324">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-325">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-325">1.0</span></span>|
|[<span data-ttu-id="fe079-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-327">ReadItem</span></span>|
|[<span data-ttu-id="fe079-328">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-329">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-330">Example</span></span>

<span data-ttu-id="fe079-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="fe079-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="fe079-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fe079-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fe079-334">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="fe079-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fe079-335">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-336">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-336">Type:</span></span>

*   [<span data-ttu-id="fe079-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fe079-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fe079-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-338">Requirements</span></span>

|<span data-ttu-id="fe079-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-339">Requirement</span></span>| <span data-ttu-id="fe079-340">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-341">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-342">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-342">1.0</span></span>|
|[<span data-ttu-id="fe079-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-344">ReadItem</span></span>|
|[<span data-ttu-id="fe079-345">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-346">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-347">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="fe079-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="fe079-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="fe079-349">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-350">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-350">Read mode</span></span>

<span data-ttu-id="fe079-351">A propriedade `location` retorna uma cadeia de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-352">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-352">Compose mode</span></span>

<span data-ttu-id="fe079-353">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-354">Type:</span></span>

*   <span data-ttu-id="fe079-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="fe079-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-356">Requirements</span></span>

|<span data-ttu-id="fe079-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-357">Requirement</span></span>| <span data-ttu-id="fe079-358">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-359">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-360">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-360">1.0</span></span>|
|[<span data-ttu-id="fe079-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-362">ReadItem</span></span>|
|[<span data-ttu-id="fe079-363">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-364">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-365">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-365">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fe079-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fe079-366">normalizedSubject :String</span></span>

<span data-ttu-id="fe079-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fe079-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject).</span><span class="sxs-lookup"><span data-stu-id="fe079-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-371">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-371">Type:</span></span>

*   <span data-ttu-id="fe079-372">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fe079-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-373">Requirements</span></span>

|<span data-ttu-id="fe079-374">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-374">Requirement</span></span>| <span data-ttu-id="fe079-375">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-376">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-377">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-377">1.0</span></span>|
|[<span data-ttu-id="fe079-378">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-379">ReadItem</span></span>|
|[<span data-ttu-id="fe079-380">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-381">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-382">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-382">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="fe079-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="fe079-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="fe079-384">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="fe079-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-385">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-385">Type:</span></span>

*   [<span data-ttu-id="fe079-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="fe079-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="fe079-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-387">Requirements</span></span>

|<span data-ttu-id="fe079-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-388">Requirement</span></span>| <span data-ttu-id="fe079-389">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-390">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-391">1.3</span><span class="sxs-lookup"><span data-stu-id="fe079-391">1.3</span></span>|
|[<span data-ttu-id="fe079-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-393">ReadItem</span></span>|
|[<span data-ttu-id="fe079-394">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-395">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="fe079-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="fe079-397">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="fe079-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fe079-398">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-399">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-399">Read mode</span></span>

<span data-ttu-id="fe079-400">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="fe079-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-401">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-401">Compose mode</span></span>

<span data-ttu-id="fe079-402">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="fe079-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-403">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-403">Type:</span></span>

*   <span data-ttu-id="fe079-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-405">Requirements</span></span>

|<span data-ttu-id="fe079-406">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-406">Requirement</span></span>| <span data-ttu-id="fe079-407">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-408">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-409">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-409">1.0</span></span>|
|[<span data-ttu-id="fe079-410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-411">ReadItem</span></span>|
|[<span data-ttu-id="fe079-412">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-413">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-414">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-414">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="fe079-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe079-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe079-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-418">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-418">Type:</span></span>

*   [<span data-ttu-id="fe079-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe079-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe079-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-420">Requirements</span></span>

|<span data-ttu-id="fe079-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-421">Requirement</span></span>| <span data-ttu-id="fe079-422">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-423">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-424">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-424">1.0</span></span>|
|[<span data-ttu-id="fe079-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-426">ReadItem</span></span>|
|[<span data-ttu-id="fe079-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-428">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-429">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="fe079-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="fe079-431">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="fe079-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fe079-432">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-433">Read mode</span></span>

<span data-ttu-id="fe079-434">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="fe079-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-435">Compose mode</span></span>

<span data-ttu-id="fe079-436">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="fe079-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-437">Type:</span></span>

*   <span data-ttu-id="fe079-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-439">Requirements</span></span>

|<span data-ttu-id="fe079-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-440">Requirement</span></span>| <span data-ttu-id="fe079-441">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-442">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-443">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-443">1.0</span></span>|
|[<span data-ttu-id="fe079-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-445">ReadItem</span></span>|
|[<span data-ttu-id="fe079-446">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-447">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-448">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="fe079-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fe079-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="fe079-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="fe079-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fe079-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o representante.</span><span class="sxs-lookup"><span data-stu-id="fe079-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-454">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="fe079-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-455">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-455">Type:</span></span>

*   [<span data-ttu-id="fe079-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fe079-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fe079-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-457">Requirements</span></span>

|<span data-ttu-id="fe079-458">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-458">Requirement</span></span>| <span data-ttu-id="fe079-459">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-460">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-461">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-461">1.0</span></span>|
|[<span data-ttu-id="fe079-462">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-463">ReadItem</span></span>|
|[<span data-ttu-id="fe079-464">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-465">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-466">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-466">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="fe079-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe079-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="fe079-468">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="fe079-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fe079-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="fe079-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-471">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-471">Read mode</span></span>

<span data-ttu-id="fe079-472">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="fe079-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-473">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-473">Compose mode</span></span>

<span data-ttu-id="fe079-474">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe079-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fe079-475">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="fe079-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-476">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-476">Type:</span></span>

*   <span data-ttu-id="fe079-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="fe079-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-478">Requirements</span></span>

|<span data-ttu-id="fe079-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-479">Requirement</span></span>| <span data-ttu-id="fe079-480">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-482">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-482">1.0</span></span>|
|[<span data-ttu-id="fe079-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-484">ReadItem</span></span>|
|[<span data-ttu-id="fe079-485">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-486">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-487">Example</span></span>

<span data-ttu-id="fe079-488">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="fe079-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="fe079-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fe079-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="fe079-490">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="fe079-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fe079-491">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="fe079-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-492">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-492">Read mode</span></span>

<span data-ttu-id="fe079-p129">A propriedade `subject` retorna uma cadeia de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="fe079-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="fe079-495">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-495">Compose mode</span></span>

<span data-ttu-id="fe079-496">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="fe079-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fe079-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-497">Type:</span></span>

*   <span data-ttu-id="fe079-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fe079-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-499">Requirements</span></span>

|<span data-ttu-id="fe079-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-500">Requirement</span></span>| <span data-ttu-id="fe079-501">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-502">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-503">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-503">1.0</span></span>|
|[<span data-ttu-id="fe079-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-505">ReadItem</span></span>|
|[<span data-ttu-id="fe079-506">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-507">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="fe079-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="fe079-509">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fe079-510">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fe079-511">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-511">Read mode</span></span>

<span data-ttu-id="fe079-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="fe079-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fe079-514">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="fe079-514">Compose mode</span></span>

<span data-ttu-id="fe079-515">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fe079-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fe079-516">Type:</span></span>

*   <span data-ttu-id="fe079-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fe079-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-518">Requirements</span></span>

|<span data-ttu-id="fe079-519">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-519">Requirement</span></span>| <span data-ttu-id="fe079-520">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-521">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-522">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-522">1.0</span></span>|
|[<span data-ttu-id="fe079-523">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-524">ReadItem</span></span>|
|[<span data-ttu-id="fe079-525">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-526">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-527">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-527">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="fe079-528">Métodos</span><span class="sxs-lookup"><span data-stu-id="fe079-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fe079-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe079-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fe079-530">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="fe079-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fe079-531">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de composição.</span><span class="sxs-lookup"><span data-stu-id="fe079-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fe079-532">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="fe079-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-533">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-533">Parameters:</span></span>

|<span data-ttu-id="fe079-534">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-534">Name</span></span>| <span data-ttu-id="fe079-535">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-535">Type</span></span>| <span data-ttu-id="fe079-536">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-536">Attributes</span></span>| <span data-ttu-id="fe079-537">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="fe079-538">String</span><span class="sxs-lookup"><span data-stu-id="fe079-538">String</span></span>||<span data-ttu-id="fe079-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fe079-541">String</span><span class="sxs-lookup"><span data-stu-id="fe079-541">String</span></span>||<span data-ttu-id="fe079-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fe079-544">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-544">Object</span></span>| <span data-ttu-id="fe079-545">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-545">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-546">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-547">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-547">Object</span></span>| <span data-ttu-id="fe079-548">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-548">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-549">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe079-550">function</span><span class="sxs-lookup"><span data-stu-id="fe079-550">function</span></span>| <span data-ttu-id="fe079-551">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-551">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-552">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe079-553">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe079-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fe079-554">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="fe079-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe079-555">Erros</span><span class="sxs-lookup"><span data-stu-id="fe079-555">Errors</span></span>

| <span data-ttu-id="fe079-556">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fe079-556">Error code</span></span> | <span data-ttu-id="fe079-557">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="fe079-558">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="fe079-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="fe079-559">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="fe079-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fe079-560">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="fe079-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-561">Requirements</span></span>

|<span data-ttu-id="fe079-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-562">Requirement</span></span>| <span data-ttu-id="fe079-563">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-565">1.1</span><span class="sxs-lookup"><span data-stu-id="fe079-565">1.1</span></span>|
|[<span data-ttu-id="fe079-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-568">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-569">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-570">Example</span></span>

```js
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fe079-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe079-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fe079-572">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fe079-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de composição. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="fe079-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fe079-576">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="fe079-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fe079-577">Se o Suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aquele que você está editando; no entanto, isso não tem suporte e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="fe079-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-578">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-578">Parameters:</span></span>

|<span data-ttu-id="fe079-579">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-579">Name</span></span>| <span data-ttu-id="fe079-580">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-580">Type</span></span>| <span data-ttu-id="fe079-581">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-581">Attributes</span></span>| <span data-ttu-id="fe079-582">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="fe079-583">String</span><span class="sxs-lookup"><span data-stu-id="fe079-583">String</span></span>||<span data-ttu-id="fe079-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fe079-586">String</span><span class="sxs-lookup"><span data-stu-id="fe079-586">String</span></span>||<span data-ttu-id="fe079-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fe079-589">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-589">Object</span></span>| <span data-ttu-id="fe079-590">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-590">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-591">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-592">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-592">Object</span></span>| <span data-ttu-id="fe079-593">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-593">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-594">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe079-595">function</span><span class="sxs-lookup"><span data-stu-id="fe079-595">function</span></span>| <span data-ttu-id="fe079-596">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-596">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-597">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe079-598">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe079-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fe079-599">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="fe079-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe079-600">Erros</span><span class="sxs-lookup"><span data-stu-id="fe079-600">Errors</span></span>

| <span data-ttu-id="fe079-601">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fe079-601">Error code</span></span> | <span data-ttu-id="fe079-602">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fe079-603">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="fe079-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-604">Requirements</span></span>

|<span data-ttu-id="fe079-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-605">Requirement</span></span>| <span data-ttu-id="fe079-606">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-607">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-608">1.1</span><span class="sxs-lookup"><span data-stu-id="fe079-608">1.1</span></span>|
|[<span data-ttu-id="fe079-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-611">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-612">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-613">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-613">Example</span></span>

<span data-ttu-id="fe079-614">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="fe079-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="fe079-615">close()</span><span class="sxs-lookup"><span data-stu-id="fe079-615">close()</span></span>

<span data-ttu-id="fe079-616">Fecha o item atual que está sendo composto.</span><span class="sxs-lookup"><span data-stu-id="fe079-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="fe079-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item tiver alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="fe079-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-619">No Outlook na Web, se o item é um compromisso e já foi salvo usando `saveAsync`, o usuário é solicitado a salvar, descartar ou cancelar mesmo se não tiver havido alterações desde que o item foi salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="fe079-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="fe079-620">No cliente do Outlook para área de trabalho, se a mensagem é uma resposta embutida, o método `close` não tem efeito.</span><span class="sxs-lookup"><span data-stu-id="fe079-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-621">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-621">Requirements</span></span>

|<span data-ttu-id="fe079-622">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-622">Requirement</span></span>| <span data-ttu-id="fe079-623">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-624">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-625">1.3</span><span class="sxs-lookup"><span data-stu-id="fe079-625">1.3</span></span>|
|[<span data-ttu-id="fe079-626">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-627">Restrito</span><span class="sxs-lookup"><span data-stu-id="fe079-627">Restricted</span></span>|
|[<span data-ttu-id="fe079-628">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-629">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="fe079-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fe079-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="fe079-631">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="fe079-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-632">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe079-633">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="fe079-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fe079-634">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fe079-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fe079-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="fe079-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-638">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-638">Parameters:</span></span>

|<span data-ttu-id="fe079-639">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-639">Name</span></span>| <span data-ttu-id="fe079-640">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-640">Type</span></span>| <span data-ttu-id="fe079-641">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fe079-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fe079-642">String &#124; Object</span></span>| |<span data-ttu-id="fe079-p139">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fe079-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fe079-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="fe079-645">**OR**</span></span><br/><span data-ttu-id="fe079-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fe079-648">String</span><span class="sxs-lookup"><span data-stu-id="fe079-648">String</span></span> | <span data-ttu-id="fe079-649">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-649">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-p141">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fe079-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fe079-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fe079-653">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-653">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-654">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="fe079-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fe079-655">String</span><span class="sxs-lookup"><span data-stu-id="fe079-655">String</span></span> | | <span data-ttu-id="fe079-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fe079-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fe079-658">String</span><span class="sxs-lookup"><span data-stu-id="fe079-658">String</span></span> | | <span data-ttu-id="fe079-659">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="fe079-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fe079-660">String</span><span class="sxs-lookup"><span data-stu-id="fe079-660">String</span></span> | | <span data-ttu-id="fe079-p143">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fe079-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fe079-663">String</span><span class="sxs-lookup"><span data-stu-id="fe079-663">String</span></span> | | <span data-ttu-id="fe079-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fe079-667">function</span><span class="sxs-lookup"><span data-stu-id="fe079-667">function</span></span> | <span data-ttu-id="fe079-668">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-668">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-669">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-670">Requirements</span></span>

|<span data-ttu-id="fe079-671">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-671">Requirement</span></span>| <span data-ttu-id="fe079-672">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-673">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-674">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-674">1.0</span></span>|
|[<span data-ttu-id="fe079-675">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-676">ReadItem</span></span>|
|[<span data-ttu-id="fe079-677">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-678">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fe079-679">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fe079-679">Examples</span></span>

<span data-ttu-id="fe079-680">O código a seguir transmite uma cadeia de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="fe079-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fe079-681">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="fe079-681">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fe079-682">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="fe079-682">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fe079-683">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="fe079-683">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="fe079-684">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fe079-684">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="fe079-685">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="fe079-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fe079-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="fe079-687">Exibe um formulário de resposta que inclui somente o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="fe079-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-688">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe079-689">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="fe079-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fe079-690">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="fe079-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fe079-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="fe079-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-694">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-694">Parameters:</span></span>

|<span data-ttu-id="fe079-695">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-695">Name</span></span>| <span data-ttu-id="fe079-696">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-696">Type</span></span>| <span data-ttu-id="fe079-697">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fe079-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="fe079-698">String &#124; Object</span></span>| | <span data-ttu-id="fe079-p146">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fe079-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fe079-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="fe079-701">**OR**</span></span><br/><span data-ttu-id="fe079-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fe079-704">String</span><span class="sxs-lookup"><span data-stu-id="fe079-704">String</span></span> | <span data-ttu-id="fe079-705">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-705">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-p148">Uma cadeia de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A cadeia de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="fe079-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fe079-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fe079-709">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-709">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-710">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="fe079-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fe079-711">String</span><span class="sxs-lookup"><span data-stu-id="fe079-711">String</span></span> | | <span data-ttu-id="fe079-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fe079-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fe079-714">String</span><span class="sxs-lookup"><span data-stu-id="fe079-714">String</span></span> | | <span data-ttu-id="fe079-715">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="fe079-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fe079-716">String</span><span class="sxs-lookup"><span data-stu-id="fe079-716">String</span></span> | | <span data-ttu-id="fe079-p150">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="fe079-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fe079-719">String</span><span class="sxs-lookup"><span data-stu-id="fe079-719">String</span></span> | | <span data-ttu-id="fe079-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fe079-723">function</span><span class="sxs-lookup"><span data-stu-id="fe079-723">function</span></span> | <span data-ttu-id="fe079-724">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-724">&lt;optional&gt;</span></span> | <span data-ttu-id="fe079-725">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-726">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-726">Requirements</span></span>

|<span data-ttu-id="fe079-727">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-727">Requirement</span></span>| <span data-ttu-id="fe079-728">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-729">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-730">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-730">1.0</span></span>|
|[<span data-ttu-id="fe079-731">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-732">ReadItem</span></span>|
|[<span data-ttu-id="fe079-733">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-734">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fe079-735">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fe079-735">Examples</span></span>

<span data-ttu-id="fe079-736">O código a seguir transmite uma cadeia de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="fe079-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fe079-737">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="fe079-737">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fe079-738">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="fe079-738">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fe079-739">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="fe079-739">Reply with a body and a file attachment.</span></span>

```js
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

<span data-ttu-id="fe079-740">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="fe079-740">Reply with a body and an item attachment.</span></span>

```js
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

<span data-ttu-id="fe079-741">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="fe079-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fe079-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="fe079-743">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fe079-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-744">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-745">Requirements</span></span>

|<span data-ttu-id="fe079-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-746">Requirement</span></span>| <span data-ttu-id="fe079-747">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-748">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-749">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-749">1.0</span></span>|
|[<span data-ttu-id="fe079-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-751">ReadItem</span></span>|
|[<span data-ttu-id="fe079-752">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-753">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-754">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-754">Returns:</span></span>

<span data-ttu-id="fe079-755">Tipo: [Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fe079-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fe079-756">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-756">Example</span></span>

<span data-ttu-id="fe079-757">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-757">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="fe079-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fe079-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fe079-759">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fe079-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-760">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-761">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-761">Parameters:</span></span>

|<span data-ttu-id="fe079-762">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-762">Name</span></span>| <span data-ttu-id="fe079-763">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-763">Type</span></span>| <span data-ttu-id="fe079-764">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="fe079-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fe079-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="fe079-766">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="fe079-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-767">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-767">Requirements</span></span>

|<span data-ttu-id="fe079-768">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-768">Requirement</span></span>| <span data-ttu-id="fe079-769">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-770">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-771">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-771">1.0</span></span>|
|[<span data-ttu-id="fe079-772">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-773">Restrito</span><span class="sxs-lookup"><span data-stu-id="fe079-773">Restricted</span></span>|
|[<span data-ttu-id="fe079-774">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-775">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-776">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-776">Returns:</span></span>

<span data-ttu-id="fe079-777">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="fe079-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fe079-778">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="fe079-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="fe079-779">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="fe079-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fe079-780">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="fe079-781">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="fe079-781">Value of `entityType`</span></span> | <span data-ttu-id="fe079-782">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="fe079-782">Type of objects in returned array</span></span> | <span data-ttu-id="fe079-783">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="fe079-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="fe079-784">String</span><span class="sxs-lookup"><span data-stu-id="fe079-784">String</span></span> | <span data-ttu-id="fe079-785">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fe079-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="fe079-786">Contato</span><span class="sxs-lookup"><span data-stu-id="fe079-786">Contact</span></span> | <span data-ttu-id="fe079-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe079-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="fe079-788">String</span><span class="sxs-lookup"><span data-stu-id="fe079-788">String</span></span> | <span data-ttu-id="fe079-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe079-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="fe079-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fe079-790">MeetingSuggestion</span></span> | <span data-ttu-id="fe079-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe079-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="fe079-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fe079-792">PhoneNumber</span></span> | <span data-ttu-id="fe079-793">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fe079-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="fe079-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fe079-794">TaskSuggestion</span></span> | <span data-ttu-id="fe079-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fe079-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="fe079-796">String</span><span class="sxs-lookup"><span data-stu-id="fe079-796">String</span></span> | <span data-ttu-id="fe079-797">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="fe079-797">**Restricted**</span></span> |

<span data-ttu-id="fe079-798">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fe079-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fe079-799">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-799">Example</span></span>

<span data-ttu-id="fe079-800">O exemplo a seguir mostra como acessar uma matriz de cadeias de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="fe079-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```js
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="fe079-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fe079-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fe079-802">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fe079-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-803">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe079-804">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="fe079-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-805">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-805">Parameters:</span></span>

|<span data-ttu-id="fe079-806">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-806">Name</span></span>| <span data-ttu-id="fe079-807">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-807">Type</span></span>| <span data-ttu-id="fe079-808">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fe079-809">String</span><span class="sxs-lookup"><span data-stu-id="fe079-809">String</span></span>|<span data-ttu-id="fe079-810">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="fe079-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-811">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-811">Requirements</span></span>

|<span data-ttu-id="fe079-812">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-812">Requirement</span></span>| <span data-ttu-id="fe079-813">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-814">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-815">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-815">1.0</span></span>|
|[<span data-ttu-id="fe079-816">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-817">ReadItem</span></span>|
|[<span data-ttu-id="fe079-818">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-819">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-820">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-820">Returns:</span></span>

<span data-ttu-id="fe079-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="fe079-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fe079-823">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fe079-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="fe079-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fe079-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fe079-825">Retorna valores de cadeia de caracteres do item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fe079-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-826">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe079-p154">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="fe079-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fe079-830">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="fe079-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fe079-831">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="fe079-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fe079-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) para recuperar todo o corpo.</span><span class="sxs-lookup"><span data-stu-id="fe079-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fe079-835">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-835">Requirements</span></span>

|<span data-ttu-id="fe079-836">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-836">Requirement</span></span>| <span data-ttu-id="fe079-837">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-838">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-839">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-839">1.0</span></span>|
|[<span data-ttu-id="fe079-840">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-841">ReadItem</span></span>|
|[<span data-ttu-id="fe079-842">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-843">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-844">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-844">Returns:</span></span>

<span data-ttu-id="fe079-p156">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="fe079-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fe079-847">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fe079-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fe079-848">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fe079-849">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-849">Example</span></span>

<span data-ttu-id="fe079-850">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="fe079-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fe079-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="fe079-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fe079-852">Retorna valores de cadeia de caracteres do item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fe079-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-853">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="fe079-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fe079-854">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="fe079-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fe079-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular como `.*` para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="fe079-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-857">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-857">Parameters:</span></span>

|<span data-ttu-id="fe079-858">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-858">Name</span></span>| <span data-ttu-id="fe079-859">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-859">Type</span></span>| <span data-ttu-id="fe079-860">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fe079-861">String</span><span class="sxs-lookup"><span data-stu-id="fe079-861">String</span></span>|<span data-ttu-id="fe079-862">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="fe079-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-863">Requirements</span></span>

|<span data-ttu-id="fe079-864">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-864">Requirement</span></span>| <span data-ttu-id="fe079-865">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-866">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-867">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-867">1.0</span></span>|
|[<span data-ttu-id="fe079-868">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-869">ReadItem</span></span>|
|[<span data-ttu-id="fe079-870">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-871">Leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-872">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-872">Returns:</span></span>

<span data-ttu-id="fe079-873">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="fe079-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fe079-874">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fe079-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fe079-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="fe079-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fe079-876">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fe079-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="fe079-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fe079-878">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fe079-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="fe079-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-881">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-881">Parameters:</span></span>

|<span data-ttu-id="fe079-882">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-882">Name</span></span>| <span data-ttu-id="fe079-883">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-883">Type</span></span>| <span data-ttu-id="fe079-884">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-884">Attributes</span></span>| <span data-ttu-id="fe079-885">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="fe079-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fe079-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fe079-p159">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="fe079-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="fe079-890">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-890">Object</span></span>| <span data-ttu-id="fe079-891">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-891">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-892">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-893">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-893">Object</span></span>| <span data-ttu-id="fe079-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-894">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-895">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe079-896">function</span><span class="sxs-lookup"><span data-stu-id="fe079-896">function</span></span>||<span data-ttu-id="fe079-897">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fe079-898">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="fe079-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fe079-899">Para acessar a propriedade de origem de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="fe079-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-900">Requirements</span></span>

|<span data-ttu-id="fe079-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-901">Requirement</span></span>| <span data-ttu-id="fe079-902">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-903">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-904">1.2</span><span class="sxs-lookup"><span data-stu-id="fe079-904">1.2</span></span>|
|[<span data-ttu-id="fe079-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-907">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-908">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fe079-909">Retorna:</span><span class="sxs-lookup"><span data-stu-id="fe079-909">Returns:</span></span>

<span data-ttu-id="fe079-910">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="fe079-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fe079-911">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="fe079-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fe079-912">String</span><span class="sxs-lookup"><span data-stu-id="fe079-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fe079-913">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-913">Example</span></span>

```js
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fe079-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fe079-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fe079-915">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="fe079-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fe079-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="fe079-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-919">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-919">Parameters:</span></span>

|<span data-ttu-id="fe079-920">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-920">Name</span></span>| <span data-ttu-id="fe079-921">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-921">Type</span></span>| <span data-ttu-id="fe079-922">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-922">Attributes</span></span>| <span data-ttu-id="fe079-923">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fe079-924">function</span><span class="sxs-lookup"><span data-stu-id="fe079-924">function</span></span>||<span data-ttu-id="fe079-925">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fe079-926">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe079-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fe079-927">Esse objeto pode ser usado para obter, definir e remover as propriedades personalizadas do item e salvar as alterações na propriedade personalizada definida de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="fe079-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="fe079-928">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-928">Object</span></span>| <span data-ttu-id="fe079-929">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-929">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-930">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="fe079-931">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-932">Requirements</span></span>

|<span data-ttu-id="fe079-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-933">Requirement</span></span>| <span data-ttu-id="fe079-934">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-935">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-936">1.0</span><span class="sxs-lookup"><span data-stu-id="fe079-936">1.0</span></span>|
|[<span data-ttu-id="fe079-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fe079-938">ReadItem</span></span>|
|[<span data-ttu-id="fe079-939">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-940">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fe079-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-941">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-941">Example</span></span>

<span data-ttu-id="fe079-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="fe079-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fe079-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fe079-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fe079-946">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="fe079-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fe079-p165">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="fe079-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-951">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-951">Parameters:</span></span>

|<span data-ttu-id="fe079-952">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-952">Name</span></span>| <span data-ttu-id="fe079-953">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-953">Type</span></span>| <span data-ttu-id="fe079-954">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-954">Attributes</span></span>| <span data-ttu-id="fe079-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="fe079-956">String</span><span class="sxs-lookup"><span data-stu-id="fe079-956">String</span></span>||<span data-ttu-id="fe079-p166">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="fe079-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="fe079-959">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-959">Object</span></span>| <span data-ttu-id="fe079-960">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-960">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-961">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-962">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-962">Object</span></span>| <span data-ttu-id="fe079-963">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-963">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-964">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fe079-965">function</span><span class="sxs-lookup"><span data-stu-id="fe079-965">function</span></span>| <span data-ttu-id="fe079-966">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-966">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-967">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fe079-968">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="fe079-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fe079-969">Erros</span><span class="sxs-lookup"><span data-stu-id="fe079-969">Errors</span></span>

| <span data-ttu-id="fe079-970">Código de erro</span><span class="sxs-lookup"><span data-stu-id="fe079-970">Error code</span></span> | <span data-ttu-id="fe079-971">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="fe079-972">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="fe079-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-973">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-973">Requirements</span></span>

|<span data-ttu-id="fe079-974">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-974">Requirement</span></span>| <span data-ttu-id="fe079-975">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-976">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-977">1.1</span><span class="sxs-lookup"><span data-stu-id="fe079-977">1.1</span></span>|
|[<span data-ttu-id="fe079-978">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-980">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-981">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-982">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-982">Example</span></span>

<span data-ttu-id="fe079-983">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="fe079-983">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="fe079-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fe079-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="fe079-985">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="fe079-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="fe079-p167">Quando chamado, este método salva a mensagem atual como um rascunho e retorna a identificação do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="fe079-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-989">Se seu suplemento chamar `saveAsync` em um item no modo de redação a fim de obter um `itemId` para usar com a API EWS ou REST, esteja ciente de que quando o Outlook está no modo de cache, pode levar alguns instantes até que o item esteja realmente sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="fe079-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="fe079-990">Até que o item esteja sincronizado, usar `itemId` retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="fe079-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="fe079-p169">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo Redigir, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="fe079-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="fe079-994">Os clientes a seguir têm diferentes comportamentos para `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="fe079-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="fe079-995">O Outlook para Mac não dá suporte a `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fe079-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="fe079-996">Chamar `saveAsync` em uma reunião no Outlook para Mac retorna um erro.</span><span class="sxs-lookup"><span data-stu-id="fe079-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="fe079-997">O Outlook na Web sempre envia um convite ou atualização quando `saveAsync` é chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="fe079-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-998">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-998">Parameters:</span></span>

|<span data-ttu-id="fe079-999">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-999">Name</span></span>| <span data-ttu-id="fe079-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-1000">Type</span></span>| <span data-ttu-id="fe079-1001">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-1001">Attributes</span></span>| <span data-ttu-id="fe079-1002">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="fe079-1003">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-1003">Object</span></span>| <span data-ttu-id="fe079-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-1005">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-1006">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-1006">Object</span></span>| <span data-ttu-id="fe079-1007">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-1008">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="fe079-1009">function</span><span class="sxs-lookup"><span data-stu-id="fe079-1009">function</span></span>||<span data-ttu-id="fe079-1010">Quando o método é concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fe079-1011">Em caso de sucesso, o identificador do item é fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="fe079-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fe079-1012">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-1012">Requirements</span></span>

|<span data-ttu-id="fe079-1013">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-1013">Requirement</span></span>| <span data-ttu-id="fe079-1014">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-1015">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="fe079-1016">1.3</span></span>|
|[<span data-ttu-id="fe079-1017">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-1019">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-1020">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fe079-1021">Exemplos</span><span class="sxs-lookup"><span data-stu-id="fe079-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="fe079-p171">A seguir apresentamos um exemplo do parâmetro `result` passado à função de retorno de chamada. A propriedade `value` contém a ID para o item.</span><span class="sxs-lookup"><span data-stu-id="fe079-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fe079-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="fe079-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fe079-1025">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fe079-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fe079-p172">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="fe079-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fe079-1029">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="fe079-1029">Parameters:</span></span>

|<span data-ttu-id="fe079-1030">Nome</span><span class="sxs-lookup"><span data-stu-id="fe079-1030">Name</span></span>| <span data-ttu-id="fe079-1031">Tipo</span><span class="sxs-lookup"><span data-stu-id="fe079-1031">Type</span></span>| <span data-ttu-id="fe079-1032">Atributos</span><span class="sxs-lookup"><span data-stu-id="fe079-1032">Attributes</span></span>| <span data-ttu-id="fe079-1033">Descrição</span><span class="sxs-lookup"><span data-stu-id="fe079-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fe079-1034">String</span><span class="sxs-lookup"><span data-stu-id="fe079-1034">String</span></span>||<span data-ttu-id="fe079-p173">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="fe079-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="fe079-1038">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-1038">Object</span></span>| <span data-ttu-id="fe079-1039">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-1040">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="fe079-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fe079-1041">Objeto</span><span class="sxs-lookup"><span data-stu-id="fe079-1041">Object</span></span>| <span data-ttu-id="fe079-1042">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-1043">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="fe079-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="fe079-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fe079-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="fe079-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="fe079-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="fe079-p174">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="fe079-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fe079-p175">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retorna um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="fe079-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fe079-1050">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, HTML será usado; se o campo for texto, texto sem formatação será usado.</span><span class="sxs-lookup"><span data-stu-id="fe079-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="fe079-1051">function</span><span class="sxs-lookup"><span data-stu-id="fe079-1051">function</span></span>||<span data-ttu-id="fe079-1052">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="fe079-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fe079-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fe079-1053">Requirements</span></span>

|<span data-ttu-id="fe079-1054">Requisito</span><span class="sxs-lookup"><span data-stu-id="fe079-1054">Requirement</span></span>| <span data-ttu-id="fe079-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="fe079-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="fe079-1056">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fe079-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fe079-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="fe079-1057">1.2</span></span>|
|[<span data-ttu-id="fe079-1058">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="fe079-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fe079-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fe079-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="fe079-1060">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fe079-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fe079-1061">Redação</span><span class="sxs-lookup"><span data-stu-id="fe079-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fe079-1062">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fe079-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```