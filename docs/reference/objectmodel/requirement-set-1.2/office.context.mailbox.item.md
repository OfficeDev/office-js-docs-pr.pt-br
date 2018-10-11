
# <a name="item"></a><span data-ttu-id="f1efd-101">item</span><span class="sxs-lookup"><span data-stu-id="f1efd-101">item</span></span>

### <span data-ttu-id="f1efd-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f1efd-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="f1efd-p102">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="f1efd-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-106">Requirements</span></span>

|<span data-ttu-id="f1efd-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-107">Requirement</span></span>| <span data-ttu-id="f1efd-108">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-109">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-110">1.0</span></span>|
|[<span data-ttu-id="f1efd-111">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-112">Restrito</span><span class="sxs-lookup"><span data-stu-id="f1efd-112">Restricted</span></span>|
|[<span data-ttu-id="f1efd-113">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-114">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="f1efd-115">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-115">Example</span></span>

<span data-ttu-id="f1efd-116">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject`  do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="f1efd-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f1efd-117">Membros</span><span class="sxs-lookup"><span data-stu-id="f1efd-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="f1efd-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f1efd-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="f1efd-p103">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-121">Certos tipos de arquivos são bloqueados pelo Outlook devido a problemas de segurança potenciais e, portanto, não são retornados.</span><span class="sxs-lookup"><span data-stu-id="f1efd-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f1efd-122">Para obter mais informações, confira [Anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="f1efd-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-123">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-123">Type:</span></span>

*   <span data-ttu-id="f1efd-124">Array. <[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f1efd-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-125">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-125">Requirements</span></span>

|<span data-ttu-id="f1efd-126">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-126">Requirement</span></span>| <span data-ttu-id="f1efd-127">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-128">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-129">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-129">1.0</span></span>|
|[<span data-ttu-id="f1efd-130">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-131">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-132">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-133">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-134">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-134">Example</span></span>

<span data-ttu-id="f1efd-135">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="f1efd-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="f1efd-137">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f1efd-138">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="f1efd-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-139">Type:</span></span>

*   [<span data-ttu-id="f1efd-140">Destinatários</span><span class="sxs-lookup"><span data-stu-id="f1efd-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f1efd-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-141">Requirements</span></span>

|<span data-ttu-id="f1efd-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-142">Requirement</span></span>| <span data-ttu-id="f1efd-143">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-144">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-145">1.1</span><span class="sxs-lookup"><span data-stu-id="f1efd-145">1.1</span></span>|
|[<span data-ttu-id="f1efd-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-147">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-148">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-149">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="f1efd-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="f1efd-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="f1efd-152">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-153">Type:</span></span>

*   [<span data-ttu-id="f1efd-154">Corpo</span><span class="sxs-lookup"><span data-stu-id="f1efd-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="f1efd-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-155">Requirements</span></span>

|<span data-ttu-id="f1efd-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-156">Requirement</span></span>| <span data-ttu-id="f1efd-157">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-158">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f1efd-159">1.1</span></span>|
|[<span data-ttu-id="f1efd-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-161">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-162">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="f1efd-164">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="f1efd-165">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f1efd-166">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-167">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-167">Read mode</span></span>

<span data-ttu-id="f1efd-p107">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-170">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-170">Compose mode</span></span>

<span data-ttu-id="f1efd-171">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-172">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-172">Type:</span></span>

*   <span data-ttu-id="f1efd-173">Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-174">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-174">Requirements</span></span>

|<span data-ttu-id="f1efd-175">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-175">Requirement</span></span>| <span data-ttu-id="f1efd-176">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-177">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-178">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-178">1.0</span></span>|
|[<span data-ttu-id="f1efd-179">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-180">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-181">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-182">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f1efd-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f1efd-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="f1efd-185">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="f1efd-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f1efd-p108">Você pode obter um inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f1efd-p109">Você obtém nulo para esta propriedade para um novo item em um formulário de redação. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-190">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-190">Type:</span></span>

*   <span data-ttu-id="f1efd-191">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-192">Requirements</span></span>

|<span data-ttu-id="f1efd-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-193">Requirement</span></span>| <span data-ttu-id="f1efd-194">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-195">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-196">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-196">1.0</span></span>|
|[<span data-ttu-id="f1efd-197">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-198">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-199">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-200">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f1efd-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="f1efd-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="f1efd-p110">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-204">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-204">Type:</span></span>

*   <span data-ttu-id="f1efd-205">Data</span><span class="sxs-lookup"><span data-stu-id="f1efd-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-206">Requirements</span></span>

|<span data-ttu-id="f1efd-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-207">Requirement</span></span>| <span data-ttu-id="f1efd-208">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-209">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-210">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-210">1.0</span></span>|
|[<span data-ttu-id="f1efd-211">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-212">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-213">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-214">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f1efd-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="f1efd-216">dateTimeModified :Date</span></span>

<span data-ttu-id="f1efd-p111">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-219">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-220">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-220">Type:</span></span>

*   <span data-ttu-id="f1efd-221">Data</span><span class="sxs-lookup"><span data-stu-id="f1efd-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-222">Requirements</span></span>

|<span data-ttu-id="f1efd-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-223">Requirement</span></span>| <span data-ttu-id="f1efd-224">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-225">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-226">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-226">1.0</span></span>|
|[<span data-ttu-id="f1efd-227">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-228">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-229">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-230">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-231">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="f1efd-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="f1efd-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="f1efd-233">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="f1efd-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f1efd-p112">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor da propriedade final para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-236">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-236">Read mode</span></span>

<span data-ttu-id="f1efd-237">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-238">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-238">Compose mode</span></span>

<span data-ttu-id="f1efd-239">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f1efd-240">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC do servidor.</span><span class="sxs-lookup"><span data-stu-id="f1efd-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-241">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-241">Type:</span></span>

*   <span data-ttu-id="f1efd-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="f1efd-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-243">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-243">Requirements</span></span>

|<span data-ttu-id="f1efd-244">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-244">Requirement</span></span>| <span data-ttu-id="f1efd-245">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-246">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-247">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-247">1.0</span></span>|
|[<span data-ttu-id="f1efd-248">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-249">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-250">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-251">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-252">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-252">Example</span></span>

<span data-ttu-id="f1efd-253">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="f1efd-254">de :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f1efd-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="f1efd-p113">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f1efd-p114">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-259">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-260">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-260">Type:</span></span>

*   [<span data-ttu-id="f1efd-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f1efd-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f1efd-262">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-262">Requirements</span></span>

|<span data-ttu-id="f1efd-263">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-263">Requirement</span></span>| <span data-ttu-id="f1efd-264">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-265">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-266">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-266">1.0</span></span>|
|[<span data-ttu-id="f1efd-267">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-268">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-269">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-270">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f1efd-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f1efd-271">internetMessageId :String</span></span>

<span data-ttu-id="f1efd-p115">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-274">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-274">Type:</span></span>

*   <span data-ttu-id="f1efd-275">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-276">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-276">Requirements</span></span>

|<span data-ttu-id="f1efd-277">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-277">Requirement</span></span>| <span data-ttu-id="f1efd-278">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-279">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-280">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-280">1.0</span></span>|
|[<span data-ttu-id="f1efd-281">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-282">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-283">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-284">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-285">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f1efd-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f1efd-286">itemClass :String</span></span>

<span data-ttu-id="f1efd-p116">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f1efd-p117">A propriedade `itemClass` especifica a classe da mensagens do item selecionado. A seguir, estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f1efd-291">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-291">Type</span></span> | <span data-ttu-id="f1efd-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-292">Description</span></span> | <span data-ttu-id="f1efd-293">classe de item</span><span class="sxs-lookup"><span data-stu-id="f1efd-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f1efd-294">Itens de compromisso</span><span class="sxs-lookup"><span data-stu-id="f1efd-294">Appointment items</span></span> | <span data-ttu-id="f1efd-295">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="f1efd-296">Itens de mensagem</span><span class="sxs-lookup"><span data-stu-id="f1efd-296">Message items</span></span> | <span data-ttu-id="f1efd-297">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="f1efd-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f1efd-298">Você pode criar classes de mensagem personalizadas que ampliam uma classe de mensagem padrão, como, por exemplo, uma classe de mensagem de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="f1efd-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-299">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-299">Type:</span></span>

*   <span data-ttu-id="f1efd-300">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-301">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-301">Requirements</span></span>

|<span data-ttu-id="f1efd-302">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-302">Requirement</span></span>| <span data-ttu-id="f1efd-303">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-304">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-305">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-305">1.0</span></span>|
|[<span data-ttu-id="f1efd-306">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-307">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-308">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-309">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-310">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f1efd-311">(anulável) itemId :String</span><span class="sxs-lookup"><span data-stu-id="f1efd-311">(nullable) itemId :String</span></span>

<span data-ttu-id="f1efd-p118">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-314">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="f1efd-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f1efd-315">A propriedade `itemId` não é idêntica à ID de entrada do Outlook ou à ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f1efd-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f1efd-316">Antes de fazer chamadas da API REST usando esse valor, ela deve ser convertida usando `Office.context.mailbox.convertToRestId`, que está disponível iniciando no conjunto de requisitos 1.3.</span><span class="sxs-lookup"><span data-stu-id="f1efd-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="f1efd-317">Para obter mais detalhes, confira [Use as APIs REST do Outlook a partir de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="f1efd-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-318">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-318">Type:</span></span>

*   <span data-ttu-id="f1efd-319">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-320">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-320">Requirements</span></span>

|<span data-ttu-id="f1efd-321">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-321">Requirement</span></span>| <span data-ttu-id="f1efd-322">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-323">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-324">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-324">1.0</span></span>|
|[<span data-ttu-id="f1efd-325">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-326">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-327">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-328">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-329">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-329">Example</span></span>

<span data-ttu-id="f1efd-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="f1efd-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f1efd-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f1efd-333">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="f1efd-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f1efd-334">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-335">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-335">Type:</span></span>

*   [<span data-ttu-id="f1efd-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f1efd-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f1efd-337">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-337">Requirements</span></span>

|<span data-ttu-id="f1efd-338">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-338">Requirement</span></span>| <span data-ttu-id="f1efd-339">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-340">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-341">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-341">1.0</span></span>|
|[<span data-ttu-id="f1efd-342">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-343">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-344">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-345">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-346">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="f1efd-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="f1efd-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="f1efd-348">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-349">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-349">Read mode</span></span>

<span data-ttu-id="f1efd-350">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-351">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-351">Compose mode</span></span>

<span data-ttu-id="f1efd-352">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-353">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-353">Type:</span></span>

*   <span data-ttu-id="f1efd-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="f1efd-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-355">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-355">Requirements</span></span>

|<span data-ttu-id="f1efd-356">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-356">Requirement</span></span>| <span data-ttu-id="f1efd-357">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-358">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-359">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-359">1.0</span></span>|
|[<span data-ttu-id="f1efd-360">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-361">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-362">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-363">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-364">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f1efd-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f1efd-365">normalizedSubject :String</span></span>

<span data-ttu-id="f1efd-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f1efd-p122">A propriedade normalizedSubject obtém o assunto do item, com quaisquer prefixos padrão (como `RE:` e `FW:`), que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject).</span><span class="sxs-lookup"><span data-stu-id="f1efd-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-370">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-370">Type:</span></span>

*   <span data-ttu-id="f1efd-371">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-372">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-372">Requirements</span></span>

|<span data-ttu-id="f1efd-373">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-373">Requirement</span></span>| <span data-ttu-id="f1efd-374">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-375">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-376">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-376">1.0</span></span>|
|[<span data-ttu-id="f1efd-377">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-378">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-379">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-380">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-381">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="f1efd-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="f1efd-383">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="f1efd-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f1efd-384">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-385">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-385">Read mode</span></span>

<span data-ttu-id="f1efd-386">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="f1efd-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-387">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-387">Compose mode</span></span>

<span data-ttu-id="f1efd-388">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f1efd-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-389">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-389">Type:</span></span>

*   <span data-ttu-id="f1efd-390">Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-391">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-391">Requirements</span></span>

|<span data-ttu-id="f1efd-392">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-392">Requirement</span></span>| <span data-ttu-id="f1efd-393">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-394">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-395">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-395">1.0</span></span>|
|[<span data-ttu-id="f1efd-396">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-397">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-398">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-399">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-400">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="f1efd-401"> organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f1efd-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="f1efd-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-404">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-404">Type:</span></span>

*   [<span data-ttu-id="f1efd-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f1efd-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f1efd-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-406">Requirements</span></span>

|<span data-ttu-id="f1efd-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-407">Requirement</span></span>| <span data-ttu-id="f1efd-408">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-409">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-410">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-410">1.0</span></span>|
|[<span data-ttu-id="f1efd-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-412">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-413">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-414">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="f1efd-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="f1efd-417">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="f1efd-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f1efd-418">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-419">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-419">Read mode</span></span>

<span data-ttu-id="f1efd-420">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="f1efd-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-421">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-421">Compose mode</span></span>

<span data-ttu-id="f1efd-422">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="f1efd-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-423">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-423">Type:</span></span>

*   <span data-ttu-id="f1efd-424">Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-425">Requirements</span></span>

|<span data-ttu-id="f1efd-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-426">Requirement</span></span>| <span data-ttu-id="f1efd-427">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-428">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-429">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-429">1.0</span></span>|
|[<span data-ttu-id="f1efd-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-431">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-432">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-433">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="f1efd-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f1efd-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="f1efd-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f1efd-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um representante. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-440">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-441">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-441">Type:</span></span>

*   [<span data-ttu-id="f1efd-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f1efd-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f1efd-443">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-443">Requirements</span></span>

|<span data-ttu-id="f1efd-444">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-444">Requirement</span></span>| <span data-ttu-id="f1efd-445">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-446">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-447">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-447">1.0</span></span>|
|[<span data-ttu-id="f1efd-448">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-449">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-450">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-451">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-452">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="f1efd-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="f1efd-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="f1efd-454">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="f1efd-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f1efd-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-457">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-457">Read mode</span></span>

<span data-ttu-id="f1efd-458">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-459">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-459">Compose mode</span></span>

<span data-ttu-id="f1efd-460">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f1efd-461">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="f1efd-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-462">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-462">Type:</span></span>

*   <span data-ttu-id="f1efd-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="f1efd-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-464">Requirements</span></span>

|<span data-ttu-id="f1efd-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-465">Requirement</span></span>| <span data-ttu-id="f1efd-466">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-467">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-468">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-468">1.0</span></span>|
|[<span data-ttu-id="f1efd-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-470">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-471">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-472">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-473">Example</span></span>

<span data-ttu-id="f1efd-474">O exemplo a seguir define a hora de início de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="f1efd-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f1efd-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="f1efd-476">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f1efd-477">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="f1efd-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-478">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-478">Read mode</span></span>

<span data-ttu-id="f1efd-p129">A propriedade `subject` retorna uma sequência de caracteres. Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto, exceto pelos prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-481">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-481">Compose mode</span></span>

<span data-ttu-id="f1efd-482">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="f1efd-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f1efd-483">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-483">Type:</span></span>

*   <span data-ttu-id="f1efd-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f1efd-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-485">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-485">Requirements</span></span>

|<span data-ttu-id="f1efd-486">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-486">Requirement</span></span>| <span data-ttu-id="f1efd-487">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-488">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-489">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-489">1.0</span></span>|
|[<span data-ttu-id="f1efd-490">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-491">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-492">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-493">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="f1efd-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="f1efd-495">Fornece acesso aos destinatários na linha **To** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f1efd-496">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f1efd-497">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-497">Read mode</span></span>

<span data-ttu-id="f1efd-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **To** da mensagem. A coleta é limitada a um número máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f1efd-500">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="f1efd-500">Compose mode</span></span>

<span data-ttu-id="f1efd-501">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha de mensagem **To**.</span><span class="sxs-lookup"><span data-stu-id="f1efd-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f1efd-502">Tipo:</span><span class="sxs-lookup"><span data-stu-id="f1efd-502">Type:</span></span>

*   <span data-ttu-id="f1efd-503">Array. <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f1efd-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-504">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-504">Requirements</span></span>

|<span data-ttu-id="f1efd-505">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-505">Requirement</span></span>| <span data-ttu-id="f1efd-506">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-507">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-508">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-508">1.0</span></span>|
|[<span data-ttu-id="f1efd-509">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-510">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-511">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-512">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-513">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f1efd-514">Métodos</span><span class="sxs-lookup"><span data-stu-id="f1efd-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f1efd-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f1efd-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f1efd-516">Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f1efd-517">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="f1efd-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f1efd-518">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f1efd-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-519">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-519">Parameters:</span></span>

|<span data-ttu-id="f1efd-520">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-520">Name</span></span>| <span data-ttu-id="f1efd-521">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-521">Type</span></span>| <span data-ttu-id="f1efd-522">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-522">Attributes</span></span>| <span data-ttu-id="f1efd-523">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f1efd-524">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-524">String</span></span>||<span data-ttu-id="f1efd-p132">O URI que fornece o local do arquivo anexado à mensagem ou compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f1efd-527">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-527">String</span></span>||<span data-ttu-id="f1efd-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f1efd-530">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-530">Object</span></span>| <span data-ttu-id="f1efd-531">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-531">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-532">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f1efd-533">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-533">Object</span></span>| <span data-ttu-id="f1efd-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-534">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-535">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f1efd-536">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-536">function</span></span>| <span data-ttu-id="f1efd-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-537">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-538">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f1efd-539">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f1efd-540">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f1efd-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f1efd-541">Erros</span><span class="sxs-lookup"><span data-stu-id="f1efd-541">Errors</span></span>

| <span data-ttu-id="f1efd-542">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f1efd-542">Error code</span></span> | <span data-ttu-id="f1efd-543">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f1efd-544">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="f1efd-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f1efd-545">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="f1efd-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f1efd-546">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f1efd-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-547">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-547">Requirements</span></span>

|<span data-ttu-id="f1efd-548">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-548">Requirement</span></span>| <span data-ttu-id="f1efd-549">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-550">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-551">1.1</span><span class="sxs-lookup"><span data-stu-id="f1efd-551">1.1</span></span>|
|[<span data-ttu-id="f1efd-552">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="f1efd-554">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-555">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-556">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f1efd-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f1efd-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f1efd-558">Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f1efd-p134">O método `addItemAttachmentAsync` anexa o item com o identificador do Exchange especificado ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método é chamado com um parâmetro, `asyncResult`, que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f1efd-562">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="f1efd-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f1efd-563">Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens a itens que não sejam aqueles que você está editando; no entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-564">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-564">Parameters:</span></span>

|<span data-ttu-id="f1efd-565">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-565">Name</span></span>| <span data-ttu-id="f1efd-566">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-566">Type</span></span>| <span data-ttu-id="f1efd-567">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-567">Attributes</span></span>| <span data-ttu-id="f1efd-568">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f1efd-569">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-569">String</span></span>||<span data-ttu-id="f1efd-p135">O identificador do Exchange do item a anexar. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f1efd-572">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-572">String</span></span>||<span data-ttu-id="f1efd-p136">O assunto do item a anexar. O tamanho máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f1efd-575">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-575">Object</span></span>| <span data-ttu-id="f1efd-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-576">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-577">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f1efd-578">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-578">Object</span></span>| <span data-ttu-id="f1efd-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-579">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-580">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f1efd-581">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-581">function</span></span>| <span data-ttu-id="f1efd-582">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-582">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-583">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f1efd-584">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f1efd-585">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornece uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="f1efd-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f1efd-586">Erros</span><span class="sxs-lookup"><span data-stu-id="f1efd-586">Errors</span></span>

| <span data-ttu-id="f1efd-587">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f1efd-587">Error code</span></span> | <span data-ttu-id="f1efd-588">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f1efd-589">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="f1efd-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-590">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-590">Requirements</span></span>

|<span data-ttu-id="f1efd-591">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-591">Requirement</span></span>| <span data-ttu-id="f1efd-592">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-593">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-594">1.1</span><span class="sxs-lookup"><span data-stu-id="f1efd-594">1.1</span></span>|
|[<span data-ttu-id="f1efd-595">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="f1efd-597">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-598">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-599">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-599">Example</span></span>

<span data-ttu-id="f1efd-600">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f1efd-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f1efd-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f1efd-602">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-603">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f1efd-604">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f1efd-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f1efd-605">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyAllForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f1efd-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f1efd-p137">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-609">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-609">Parameters:</span></span>

|<span data-ttu-id="f1efd-610">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-610">Name</span></span>| <span data-ttu-id="f1efd-611">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-611">Type</span></span>| <span data-ttu-id="f1efd-612">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f1efd-613">String | Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-613">String &#124; Object</span></span>| |<span data-ttu-id="f1efd-p138">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f1efd-616">**OU**</span><span class="sxs-lookup"><span data-stu-id="f1efd-616">**OR**</span></span><br/><span data-ttu-id="f1efd-p139">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f1efd-619">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-619">String</span></span> | <span data-ttu-id="f1efd-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-620">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-p140">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f1efd-623">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f1efd-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-624">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-625">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f1efd-626">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-626">String</span></span> | | <span data-ttu-id="f1efd-p141">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f1efd-629">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-629">String</span></span> | | <span data-ttu-id="f1efd-630">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f1efd-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f1efd-631">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-631">String</span></span> | | <span data-ttu-id="f1efd-p142">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f1efd-634">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-634">String</span></span> | | <span data-ttu-id="f1efd-p143">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f1efd-638">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-638">function</span></span> | <span data-ttu-id="f1efd-639">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-639">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-640">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-641">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-641">Requirements</span></span>

|<span data-ttu-id="f1efd-642">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-642">Requirement</span></span>| <span data-ttu-id="f1efd-643">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-644">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-645">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-645">1.0</span></span>|
|[<span data-ttu-id="f1efd-646">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-647">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-648">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-649">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f1efd-650">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f1efd-650">Examples</span></span>

<span data-ttu-id="f1efd-651">O código a seguir transmite uma sequência de caracteres à função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f1efd-652">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f1efd-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f1efd-653">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f1efd-654">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-654">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f1efd-655">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-655">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f1efd-656">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f1efd-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f1efd-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="f1efd-658">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-659">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f1efd-660">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="f1efd-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f1efd-661">Se qualquer dos parâmetros da cadeia de caracteres exceder seu limite, `displayReplyForm` gera uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f1efd-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f1efd-p144">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-665">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-665">Parameters:</span></span>

|<span data-ttu-id="f1efd-666">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-666">Name</span></span>| <span data-ttu-id="f1efd-667">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-667">Type</span></span>| <span data-ttu-id="f1efd-668">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f1efd-669">String | Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-669">String &#124; Object</span></span>| | <span data-ttu-id="f1efd-p145">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f1efd-672">**OU**</span><span class="sxs-lookup"><span data-stu-id="f1efd-672">**OR**</span></span><br/><span data-ttu-id="f1efd-p146">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da maneira a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f1efd-675">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-675">String</span></span> | <span data-ttu-id="f1efd-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-676">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-p147">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f1efd-679">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f1efd-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-680">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-681">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f1efd-682">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-682">String</span></span> | | <span data-ttu-id="f1efd-p148">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f1efd-685">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-685">String</span></span> | | <span data-ttu-id="f1efd-686">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="f1efd-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f1efd-687">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-687">String</span></span> | | <span data-ttu-id="f1efd-p149">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f1efd-690">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-690">String</span></span> | | <span data-ttu-id="f1efd-p150">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f1efd-694">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-694">function</span></span> | <span data-ttu-id="f1efd-695">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-695">&lt;optional&gt;</span></span> | <span data-ttu-id="f1efd-696">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-697">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-697">Requirements</span></span>

|<span data-ttu-id="f1efd-698">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-698">Requirement</span></span>| <span data-ttu-id="f1efd-699">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-700">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-701">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-701">1.0</span></span>|
|[<span data-ttu-id="f1efd-702">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-703">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-704">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-705">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f1efd-706">Exemplos</span><span class="sxs-lookup"><span data-stu-id="f1efd-706">Examples</span></span>

<span data-ttu-id="f1efd-707">O código a seguir transmite uma sequência de caracteres à função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f1efd-708">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="f1efd-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f1efd-709">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f1efd-710">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-710">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f1efd-711">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="f1efd-711">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f1efd-712">Responder com um corpo, um anexo de arquivo, um anexo do item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="f1efd-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f1efd-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="f1efd-714">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-715">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-716">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-716">Requirements</span></span>

|<span data-ttu-id="f1efd-717">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-717">Requirement</span></span>| <span data-ttu-id="f1efd-718">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-719">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-720">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-720">1.0</span></span>|
|[<span data-ttu-id="f1efd-721">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-722">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-723">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-724">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-725">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-725">Returns:</span></span>

<span data-ttu-id="f1efd-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f1efd-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f1efd-727">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-727">Example</span></span>

<span data-ttu-id="f1efd-728">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-728">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="f1efd-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f1efd-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f1efd-730">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontrado no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-731">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-732">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-732">Parameters:</span></span>

|<span data-ttu-id="f1efd-733">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-733">Name</span></span>| <span data-ttu-id="f1efd-734">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-734">Type</span></span>| <span data-ttu-id="f1efd-735">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f1efd-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f1efd-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="f1efd-737">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="f1efd-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1efd-738">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-738">Requirements</span></span>

|<span data-ttu-id="f1efd-739">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-739">Requirement</span></span>| <span data-ttu-id="f1efd-740">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-741">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-742">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-742">1.0</span></span>|
|[<span data-ttu-id="f1efd-743">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-744">Restrito</span><span class="sxs-lookup"><span data-stu-id="f1efd-744">Restricted</span></span>|
|[<span data-ttu-id="f1efd-745">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-746">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-747">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-747">Returns:</span></span>

<span data-ttu-id="f1efd-748">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retorna nulo.</span><span class="sxs-lookup"><span data-stu-id="f1efd-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f1efd-749">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f1efd-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="f1efd-750">Caso contrário, o tipo de objetos na matriz retornada depende do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f1efd-751">Enquanto o nível de permissão mínimo a usar esse método é **Restricted**, alguns tipos de entidade requerem **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f1efd-752">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="f1efd-752">Value of `entityType`</span></span> | <span data-ttu-id="f1efd-753">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="f1efd-753">Type of objects in returned array</span></span> | <span data-ttu-id="f1efd-754">Nível de permissão exigido</span><span class="sxs-lookup"><span data-stu-id="f1efd-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f1efd-755">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-755">String</span></span> | <span data-ttu-id="f1efd-756">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f1efd-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f1efd-757">Contato</span><span class="sxs-lookup"><span data-stu-id="f1efd-757">Contact</span></span> | <span data-ttu-id="f1efd-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f1efd-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f1efd-759">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-759">String</span></span> | <span data-ttu-id="f1efd-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f1efd-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f1efd-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f1efd-761">MeetingSuggestion</span></span> | <span data-ttu-id="f1efd-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f1efd-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f1efd-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f1efd-763">PhoneNumber</span></span> | <span data-ttu-id="f1efd-764">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f1efd-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f1efd-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f1efd-765">TaskSuggestion</span></span> | <span data-ttu-id="f1efd-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f1efd-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f1efd-767">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-767">String</span></span> | <span data-ttu-id="f1efd-768">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="f1efd-768">**Restricted**</span></span> |

<span data-ttu-id="f1efd-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f1efd-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f1efd-770">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-770">Example</span></span>

<span data-ttu-id="f1efd-771">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="f1efd-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="f1efd-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f1efd-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f1efd-773">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-774">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f1efd-775">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-776">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-776">Parameters:</span></span>

|<span data-ttu-id="f1efd-777">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-777">Name</span></span>| <span data-ttu-id="f1efd-778">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-778">Type</span></span>| <span data-ttu-id="f1efd-779">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f1efd-780">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-780">String</span></span>|<span data-ttu-id="f1efd-781">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="f1efd-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1efd-782">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-782">Requirements</span></span>

|<span data-ttu-id="f1efd-783">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-783">Requirement</span></span>| <span data-ttu-id="f1efd-784">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-785">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-786">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-786">1.0</span></span>|
|[<span data-ttu-id="f1efd-787">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-788">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-789">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-790">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-791">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-791">Returns:</span></span>

<span data-ttu-id="f1efd-p152">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retorna `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retorna uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f1efd-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f1efd-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f1efd-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f1efd-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f1efd-796">Retorna valores de sequência de caracteres ao item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-797">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f1efd-p153">O método `getRegExMatches` retorna as cadeias de caracteres que correspondem à expressão regular definida em cada elemento de regra `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma cadeia de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f1efd-801">Por exemplo, considere que um manifesto de suplemento tem o seguinte elemento `Rule`:</span><span class="sxs-lookup"><span data-stu-id="f1efd-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f1efd-802">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="f1efd-p154">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1efd-805">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-805">Requirements</span></span>

|<span data-ttu-id="f1efd-806">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-806">Requirement</span></span>| <span data-ttu-id="f1efd-807">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-808">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-809">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-809">1.0</span></span>|
|[<span data-ttu-id="f1efd-810">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-811">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-812">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-813">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-814">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-814">Returns:</span></span>

<span data-ttu-id="f1efd-p155">Um objeto que contém matrizes de cadeias de caracteres que correspondem às expressões regulares definidas no arquivo XML do manifesto. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f1efd-817">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f1efd-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f1efd-818">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f1efd-819">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-819">Example</span></span>

<span data-ttu-id="f1efd-820">O exemplo a seguir mostra como acessar a matriz de correspondências para os elementos <rule> da expressão regular, `fruits` e `veggies`, que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="f1efd-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f1efd-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="f1efd-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f1efd-822">Retorna valores de cadeia de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f1efd-823">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="f1efd-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f1efd-824">O método `getRegExMatchesByName` retorna as cadeias de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f1efd-p156">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar mais o corpo e não tentar retornar todo o corpo do item. Usar uma expressão regular, como `.*`, para obter todo o corpo de um item nem sempre retorna os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-827">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-827">Parameters:</span></span>

|<span data-ttu-id="f1efd-828">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-828">Name</span></span>| <span data-ttu-id="f1efd-829">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-829">Type</span></span>| <span data-ttu-id="f1efd-830">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f1efd-831">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-831">String</span></span>|<span data-ttu-id="f1efd-832">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a corresponder.</span><span class="sxs-lookup"><span data-stu-id="f1efd-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1efd-833">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-833">Requirements</span></span>

|<span data-ttu-id="f1efd-834">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-834">Requirement</span></span>| <span data-ttu-id="f1efd-835">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-836">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-837">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-837">1.0</span></span>|
|[<span data-ttu-id="f1efd-838">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-839">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-840">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-841">Leitura</span><span class="sxs-lookup"><span data-stu-id="f1efd-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-842">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-842">Returns:</span></span>

<span data-ttu-id="f1efd-843">Uma matriz que contém as cadeias de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f1efd-844">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f1efd-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f1efd-845">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f1efd-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f1efd-846">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f1efd-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f1efd-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f1efd-848">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f1efd-p157">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retorna nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retorna o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-851">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-851">Parameters:</span></span>

|<span data-ttu-id="f1efd-852">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-852">Name</span></span>| <span data-ttu-id="f1efd-853">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-853">Type</span></span>| <span data-ttu-id="f1efd-854">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-854">Attributes</span></span>| <span data-ttu-id="f1efd-855">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f1efd-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f1efd-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f1efd-p158">Solicita um formato para os dados. Se Text, o método retorna o texto sem formatação como uma cadeia de caracteres, removendo quaisquer marcas HTML presentes. Se HTML, o método retorna o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f1efd-860">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-860">Object</span></span>| <span data-ttu-id="f1efd-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-861">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-862">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f1efd-863">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-863">Object</span></span>| <span data-ttu-id="f1efd-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-864">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-865">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f1efd-866">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-866">function</span></span>||<span data-ttu-id="f1efd-867">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f1efd-868">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f1efd-869">Para acessar a propriedade original de que a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1efd-870">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-870">Requirements</span></span>

|<span data-ttu-id="f1efd-871">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-871">Requirement</span></span>| <span data-ttu-id="f1efd-872">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-873">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-874">1.2</span><span class="sxs-lookup"><span data-stu-id="f1efd-874">1.2</span></span>|
|[<span data-ttu-id="f1efd-875">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="f1efd-877">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-878">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f1efd-879">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f1efd-879">Returns:</span></span>

<span data-ttu-id="f1efd-880">Os dados selecionados como uma cadeia de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f1efd-881">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="f1efd-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f1efd-882">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f1efd-883">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-883">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f1efd-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f1efd-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f1efd-885">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="f1efd-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f1efd-p160">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retorna um objeto `CustomProperties` no retorno de chamada, que oferece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. Propriedades personalizadas não são criptografadas no item, portanto não devem ser usadas como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-889">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-889">Parameters:</span></span>

|<span data-ttu-id="f1efd-890">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-890">Name</span></span>| <span data-ttu-id="f1efd-891">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-891">Type</span></span>| <span data-ttu-id="f1efd-892">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-892">Attributes</span></span>| <span data-ttu-id="f1efd-893">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f1efd-894">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-894">function</span></span>||<span data-ttu-id="f1efd-895">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f1efd-896">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f1efd-897">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="f1efd-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f1efd-898">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-898">Object</span></span>| <span data-ttu-id="f1efd-899">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-899">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-900">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="f1efd-901">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1efd-902">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-902">Requirements</span></span>

|<span data-ttu-id="f1efd-903">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-903">Requirement</span></span>| <span data-ttu-id="f1efd-904">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-905">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-906">1.0</span><span class="sxs-lookup"><span data-stu-id="f1efd-906">1.0</span></span>|
|[<span data-ttu-id="f1efd-907">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-908">ReadItem</span></span>|
|[<span data-ttu-id="f1efd-909">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-910">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="f1efd-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-911">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-911">Example</span></span>

<span data-ttu-id="f1efd-p163">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta no servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar na propriedade personalizada `otherProp` e, então, chama o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f1efd-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f1efd-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f1efd-916">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="f1efd-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f1efd-p164">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou se o usuário começa a compor em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-921">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-921">Parameters:</span></span>

|<span data-ttu-id="f1efd-922">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-922">Name</span></span>| <span data-ttu-id="f1efd-923">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-923">Type</span></span>| <span data-ttu-id="f1efd-924">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-924">Attributes</span></span>| <span data-ttu-id="f1efd-925">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f1efd-926">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-926">String</span></span>||<span data-ttu-id="f1efd-p165">O identificador do anexo a remover. O comprimento máximo da cadeia é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="f1efd-929">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-929">Object</span></span>| <span data-ttu-id="f1efd-930">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-930">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-931">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f1efd-932">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-932">Object</span></span>| <span data-ttu-id="f1efd-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-933">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-934">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f1efd-935">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-935">function</span></span>| <span data-ttu-id="f1efd-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-936">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-937">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f1efd-938">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="f1efd-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f1efd-939">Erros</span><span class="sxs-lookup"><span data-stu-id="f1efd-939">Errors</span></span>

| <span data-ttu-id="f1efd-940">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f1efd-940">Error code</span></span> | <span data-ttu-id="f1efd-941">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f1efd-942">O identificador de anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="f1efd-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-943">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-943">Requirements</span></span>

|<span data-ttu-id="f1efd-944">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-944">Requirement</span></span>| <span data-ttu-id="f1efd-945">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-946">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-947">1.1</span><span class="sxs-lookup"><span data-stu-id="f1efd-947">1.1</span></span>|
|[<span data-ttu-id="f1efd-948">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="f1efd-950">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-951">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-952">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-952">Example</span></span>

<span data-ttu-id="f1efd-953">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="f1efd-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f1efd-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f1efd-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f1efd-955">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f1efd-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f1efd-p166">O método `setSelectedDataAsync` insere a cadeia de caracteres especificada no local do cursor no corpo ou assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou assunto, um erro será retornado. Após a inserção, o cursor é colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f1efd-959">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f1efd-959">Parameters:</span></span>

|<span data-ttu-id="f1efd-960">Nome</span><span class="sxs-lookup"><span data-stu-id="f1efd-960">Name</span></span>| <span data-ttu-id="f1efd-961">Tipo</span><span class="sxs-lookup"><span data-stu-id="f1efd-961">Type</span></span>| <span data-ttu-id="f1efd-962">Atributos</span><span class="sxs-lookup"><span data-stu-id="f1efd-962">Attributes</span></span>| <span data-ttu-id="f1efd-963">Descrição</span><span class="sxs-lookup"><span data-stu-id="f1efd-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f1efd-964">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="f1efd-964">String</span></span>||<span data-ttu-id="f1efd-p167">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 de caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f1efd-968">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-968">Object</span></span>| <span data-ttu-id="f1efd-969">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-969">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-970">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="f1efd-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f1efd-971">Object</span><span class="sxs-lookup"><span data-stu-id="f1efd-971">Object</span></span>| <span data-ttu-id="f1efd-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-972">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-973">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="f1efd-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="f1efd-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f1efd-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="f1efd-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f1efd-975">&lt;optional&gt;</span></span>|<span data-ttu-id="f1efd-p168">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f1efd-p169">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retornará um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="f1efd-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f1efd-980">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="f1efd-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f1efd-981">function</span><span class="sxs-lookup"><span data-stu-id="f1efd-981">function</span></span>||<span data-ttu-id="f1efd-982">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f1efd-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f1efd-983">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1efd-983">Requirements</span></span>

|<span data-ttu-id="f1efd-984">Requisito</span><span class="sxs-lookup"><span data-stu-id="f1efd-984">Requirement</span></span>| <span data-ttu-id="f1efd-985">Valor</span><span class="sxs-lookup"><span data-stu-id="f1efd-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1efd-986">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f1efd-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f1efd-987">1.2</span><span class="sxs-lookup"><span data-stu-id="f1efd-987">1.2</span></span>|
|[<span data-ttu-id="f1efd-988">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f1efd-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f1efd-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f1efd-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="f1efd-990">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="f1efd-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f1efd-991">Redigir</span><span class="sxs-lookup"><span data-stu-id="f1efd-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f1efd-992">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f1efd-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```