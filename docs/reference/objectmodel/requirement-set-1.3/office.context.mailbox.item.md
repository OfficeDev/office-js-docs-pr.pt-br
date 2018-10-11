
# <a name="item"></a><span data-ttu-id="8c13f-101">item</span><span class="sxs-lookup"><span data-stu-id="8c13f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="8c13f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="8c13f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="8c13f-p101">O namespace `item` é usado para acessar a mensagem, a solicitação de reunião ou o compromisso selecionado no momento. Você pode determinar o tipo de `item` usando a propriedade [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="8c13f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-105">Requirements</span></span>

|<span data-ttu-id="8c13f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-106">Requirement</span></span>| <span data-ttu-id="8c13f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-109">1.0</span></span>|
|[<span data-ttu-id="8c13f-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="8c13f-111">Restricted</span></span>|
|[<span data-ttu-id="8c13f-112">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="8c13f-114">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-114">Example</span></span>

<span data-ttu-id="8c13f-115">O exemplo de código JavaScript a seguir mostra como acessar a propriedade `subject` do item atual no Outlook.</span><span class="sxs-lookup"><span data-stu-id="8c13f-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="8c13f-116">Membros</span><span class="sxs-lookup"><span data-stu-id="8c13f-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="8c13f-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8c13f-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="8c13f-p102">Obtém uma matriz de anexos para o item. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-120">Certos tipos de arquivos são bloqueados pelo Outlook devido a problemas de segurança potenciais e portanto não são retornados.</span><span class="sxs-lookup"><span data-stu-id="8c13f-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8c13f-121">Para obter mais informações, consulte [anexos bloqueados no Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="8c13f-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-122">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-122">Type:</span></span>

*   <span data-ttu-id="8c13f-123">Array. <[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8c13f-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-124">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-124">Requirements</span></span>

|<span data-ttu-id="8c13f-125">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-125">Requirement</span></span>| <span data-ttu-id="8c13f-126">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-127">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-128">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-128">1.0</span></span>|
|[<span data-ttu-id="8c13f-129">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-130">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-131">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-133">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-133">Example</span></span>

<span data-ttu-id="8c13f-134">O código a seguir cria uma sequência de caracteres HTML com detalhes de todos os anexos no item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="8c13f-135">bcc:[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="8c13f-136">Obtém um objeto que fornece os métodos para obter ou atualizar os destinatários na linha Cco (com cópia oculta) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8c13f-137">Somente modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8c13f-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-138">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-138">Type:</span></span>

*   [<span data-ttu-id="8c13f-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="8c13f-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8c13f-140">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-140">Requirements</span></span>

|<span data-ttu-id="8c13f-141">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-141">Requirement</span></span>| <span data-ttu-id="8c13f-142">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-143">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8c13f-144">1.1</span></span>|
|[<span data-ttu-id="8c13f-145">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-146">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-147">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-148">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-149">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="8c13f-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="8c13f-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="8c13f-151">Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-152">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-152">Type:</span></span>

*   [<span data-ttu-id="8c13f-153">Body</span><span class="sxs-lookup"><span data-stu-id="8c13f-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="8c13f-154">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-154">Requirements</span></span>

|<span data-ttu-id="8c13f-155">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-155">Requirement</span></span>| <span data-ttu-id="8c13f-156">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-157">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-158">1.1</span><span class="sxs-lookup"><span data-stu-id="8c13f-158">1.1</span></span>|
|[<span data-ttu-id="8c13f-159">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-160">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-161">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-162">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="8c13f-163">cc: Array. <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="8c13f-164">Fornece acesso aos destinatários Cc (com cópia) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8c13f-165">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-166">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-166">Read mode</span></span>

<span data-ttu-id="8c13f-p106">A propriedade `cc` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Cc** da mensagem. O conjunto está limitado a um máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-169">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-169">Compose mode</span></span>

<span data-ttu-id="8c13f-170">A propriedade `cc` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha **Cc** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-171">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-171">Type:</span></span>

*   <span data-ttu-id="8c13f-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-173">Requirements</span></span>

|<span data-ttu-id="8c13f-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-174">Requirement</span></span>| <span data-ttu-id="8c13f-175">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-176">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-177">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-177">1.0</span></span>|
|[<span data-ttu-id="8c13f-178">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-179">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-180">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-181">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-182">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8c13f-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8c13f-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="8c13f-184">Obtém um identificador da conversa de email que contém uma mensagem específica.</span><span class="sxs-lookup"><span data-stu-id="8c13f-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8c13f-p107">Você pode obter um inteiro para esta propriedade se o aplicativo de email estiver ativado nos formulários de leitura ou nas respostas dos formulários de redação. Se posteriormente o usuário alterar o assunto da mensagem de resposta, ao enviar a resposta, a ID da conversa daquela mensagem será alterada e o valor obtido anteriormente não mais se aplicará.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8c13f-p108">Você obtém nulo nessa propriedade para um novo item em um formulário de redação. Se o usuário definir um assunto e salvar o item, a propriedade `conversationId` retornará um valor.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-189">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-189">Type:</span></span>

*   <span data-ttu-id="8c13f-190">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-191">Requirements</span></span>

|<span data-ttu-id="8c13f-192">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-192">Requirement</span></span>| <span data-ttu-id="8c13f-193">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-194">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-195">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-195">1.0</span></span>|
|[<span data-ttu-id="8c13f-196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-197">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-198">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-199">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8c13f-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8c13f-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="8c13f-p109">Obtém a data e a hora em que um item foi criado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-203">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-203">Type:</span></span>

*   <span data-ttu-id="8c13f-204">Date</span><span class="sxs-lookup"><span data-stu-id="8c13f-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-205">Requirements</span></span>

|<span data-ttu-id="8c13f-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-206">Requirement</span></span>| <span data-ttu-id="8c13f-207">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-208">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-209">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-209">1.0</span></span>|
|[<span data-ttu-id="8c13f-210">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-211">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-212">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-213">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8c13f-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8c13f-215">dateTimeModified :Date</span></span>

<span data-ttu-id="8c13f-p110">Obtém a data e a hora em que um item foi alterado pela última vez. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-218">Esse membro não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-219">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-219">Type:</span></span>

*   <span data-ttu-id="8c13f-220">Date</span><span class="sxs-lookup"><span data-stu-id="8c13f-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-221">Requirements</span></span>

|<span data-ttu-id="8c13f-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-222">Requirement</span></span>| <span data-ttu-id="8c13f-223">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-224">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-225">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-225">1.0</span></span>|
|[<span data-ttu-id="8c13f-226">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-227">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-228">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-229">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-230">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="8c13f-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c13f-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="8c13f-232">Obtém ou define a data e a hora em que o compromisso deve terminar.</span><span class="sxs-lookup"><span data-stu-id="8c13f-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8c13f-p111">A propriedade `end` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) para converter o valor da propriedade end para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-235">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-235">Read mode</span></span>

<span data-ttu-id="8c13f-236">A propriedade `end` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-237">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-237">Compose mode</span></span>

<span data-ttu-id="8c13f-238">A propriedade `end` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8c13f-239">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para definir a hora de término, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC do servidor.</span><span class="sxs-lookup"><span data-stu-id="8c13f-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-240">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-240">Type:</span></span>

*   <span data-ttu-id="8c13f-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c13f-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-242">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-242">Requirements</span></span>

|<span data-ttu-id="8c13f-243">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-243">Requirement</span></span>| <span data-ttu-id="8c13f-244">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-245">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-246">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-246">1.0</span></span>|
|[<span data-ttu-id="8c13f-247">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-248">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-249">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-250">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-251">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-251">Example</span></span>

<span data-ttu-id="8c13f-252">O exemplo a seguir define a hora de término de um compromisso no modo de composição usando o método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="8c13f-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c13f-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c13f-p112">Obtém o endereço de email do remetente de uma mensagem. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8c13f-p113">As propriedades `from` e [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-258">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `from` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-259">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-259">Type:</span></span>

*   [<span data-ttu-id="8c13f-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c13f-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c13f-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-261">Requirements</span></span>

|<span data-ttu-id="8c13f-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-262">Requirement</span></span>| <span data-ttu-id="8c13f-263">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-264">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-265">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-265">1.0</span></span>|
|[<span data-ttu-id="8c13f-266">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-267">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-268">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-269">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8c13f-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8c13f-270">internetMessageId :String</span></span>

<span data-ttu-id="8c13f-p114">Obtém o identificador de mensagem de Internet para uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-273">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-273">Type:</span></span>

*   <span data-ttu-id="8c13f-274">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-275">Requirements</span></span>

|<span data-ttu-id="8c13f-276">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-276">Requirement</span></span>| <span data-ttu-id="8c13f-277">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-278">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-279">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-279">1.0</span></span>|
|[<span data-ttu-id="8c13f-280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-281">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-282">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-283">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-284">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8c13f-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8c13f-285">itemClass :String</span></span>

<span data-ttu-id="8c13f-p115">Obtém a classe do item dos Serviços Web do Exchange do item selecionado. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8c13f-p116">A propriedade `itemClass` especifica a classe de mensagens do item selecionado. A seguir estão as classes de mensagem padrão para o item de mensagem ou de compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8c13f-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-290">Type</span></span> | <span data-ttu-id="8c13f-291">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-291">Description</span></span> | <span data-ttu-id="8c13f-292">classe de item</span><span class="sxs-lookup"><span data-stu-id="8c13f-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8c13f-293">Appointment items</span><span class="sxs-lookup"><span data-stu-id="8c13f-293">Appointment items</span></span> | <span data-ttu-id="8c13f-294">Esses são itens de calendário da classe de item `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="8c13f-295">Message items</span><span class="sxs-lookup"><span data-stu-id="8c13f-295">Message items</span></span> | <span data-ttu-id="8c13f-296">Incluem mensagens de email que têm a classe de mensagem padrão `IPM.Note` e solicitações de reunião, respostas e cancelamentos que utilizam `IPM.Schedule.Meeting` como a classe de mensagem básica.</span><span class="sxs-lookup"><span data-stu-id="8c13f-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8c13f-297">Você pode criar classes de mensagens personalizadas que estendem uma classe de mensagens padrão, por exemplo, uma classe de mensagens de compromisso personalizada `IPM.Appointment.Contoso` .</span><span class="sxs-lookup"><span data-stu-id="8c13f-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-298">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-298">Type:</span></span>

*   <span data-ttu-id="8c13f-299">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-300">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-300">Requirements</span></span>

|<span data-ttu-id="8c13f-301">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-301">Requirement</span></span>| <span data-ttu-id="8c13f-302">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-303">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-304">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-304">1.0</span></span>|
|[<span data-ttu-id="8c13f-305">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-306">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-307">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-308">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-309">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8c13f-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8c13f-310">(nullable) itemId :String</span></span>

<span data-ttu-id="8c13f-p117">Obtém o identificador do item dos Serviços Web do Exchange para o item atual. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-313">O identificador retornado pela propriedade `itemId` é o mesmo que o identificador do item dos Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="8c13f-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8c13f-314">A propriedade `itemId` não é idêntica à ID de entrada do Outlook ou à ID usada pela API REST do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8c13f-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8c13f-315">Antes de fazer chamadas de API REST usando esse valor, ele deve ser convertido usando [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="8c13f-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="8c13f-316">Para obter mais detalhes, consulte [Use as APIs do restante do Outlook de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="8c13f-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="8c13f-p119">A propriedade `itemId` não está disponível no modo de redação. Se for obrigatório um identificador de um item, pode ser usado o método [`saveAsync`](#saveasyncoptions-callback) para salvar o item no repositório, o que retornará o identificador do item no parâmetro [`AsyncResult.value`](/javascript/api/office/office.asyncresult) na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-319">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-319">Type:</span></span>

*   <span data-ttu-id="8c13f-320">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-321">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-321">Requirements</span></span>

|<span data-ttu-id="8c13f-322">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-322">Requirement</span></span>| <span data-ttu-id="8c13f-323">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-324">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-325">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-325">1.0</span></span>|
|[<span data-ttu-id="8c13f-326">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-327">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-328">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-329">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-330">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-330">Example</span></span>

<span data-ttu-id="8c13f-p120">O código a seguir verifica a presença de um identificador de item. Se a propriedade `itemId` retorna `null` ou `undefined`, ele salva o item no repositório e obtém o identificador do item do resultado assíncrono.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="8c13f-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8c13f-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8c13f-334">Obtém o tipo de item que representa uma instância.</span><span class="sxs-lookup"><span data-stu-id="8c13f-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8c13f-335">A propriedade `itemType` retorna um dos valores de enumeração `ItemType`, indicando se a instância do objeto `item` é uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-336">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-336">Type:</span></span>

*   [<span data-ttu-id="8c13f-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8c13f-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8c13f-338">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-338">Requirements</span></span>

|<span data-ttu-id="8c13f-339">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-339">Requirement</span></span>| <span data-ttu-id="8c13f-340">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-341">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-342">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-342">1.0</span></span>|
|[<span data-ttu-id="8c13f-343">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-344">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-345">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-346">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-347">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="8c13f-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="8c13f-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="8c13f-349">Obtém ou define o local de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-350">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-350">Read mode</span></span>

<span data-ttu-id="8c13f-351">A propriedade `location` retorna uma sequência de caracteres que contém o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-352">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-352">Compose mode</span></span>

<span data-ttu-id="8c13f-353">A propriedade `location` retorna um objeto `Location` que fornece métodos usados para obter e definir o local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-354">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-354">Type:</span></span>

*   <span data-ttu-id="8c13f-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="8c13f-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-356">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-356">Requirements</span></span>

|<span data-ttu-id="8c13f-357">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-357">Requirement</span></span>| <span data-ttu-id="8c13f-358">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-359">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-360">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-360">1.0</span></span>|
|[<span data-ttu-id="8c13f-361">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-362">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-363">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-364">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-365">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8c13f-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8c13f-366">normalizedSubject :String</span></span>

<span data-ttu-id="8c13f-p121">Obtém o assunto de um item, com todos os prefixos removidos (incluindo `RE:` e `FWD:`). Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8c13f-p122">A propriedade normalizedSubject obtém o assunto do item com quaisquer prefixos padrão (como `RE:` e `FW:`) que são adicionados por programas de email. Para obter o assunto do item com os prefixos intactos, use a propriedade [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject).</span><span class="sxs-lookup"><span data-stu-id="8c13f-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-371">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-371">Type:</span></span>

*   <span data-ttu-id="8c13f-372">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-373">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-373">Requirements</span></span>

|<span data-ttu-id="8c13f-374">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-374">Requirement</span></span>| <span data-ttu-id="8c13f-375">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-376">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-377">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-377">1.0</span></span>|
|[<span data-ttu-id="8c13f-378">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-379">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-380">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-381">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-382">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="8c13f-383">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="8c13f-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="8c13f-384">Obtém as mensagens de notificação de um item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-385">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-385">Type:</span></span>

*   [<span data-ttu-id="8c13f-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="8c13f-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="8c13f-387">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-387">Requirements</span></span>

|<span data-ttu-id="8c13f-388">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-388">Requirement</span></span>| <span data-ttu-id="8c13f-389">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-390">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-391">1.3</span><span class="sxs-lookup"><span data-stu-id="8c13f-391">1.3</span></span>|
|[<span data-ttu-id="8c13f-392">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-393">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-394">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-395">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="8c13f-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="8c13f-397">Fornece acesso aos participantes opcionais de um evento.</span><span class="sxs-lookup"><span data-stu-id="8c13f-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8c13f-398">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-399">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-399">Read mode</span></span>

<span data-ttu-id="8c13f-400">A propriedade `optionalAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante opcional da reunião.</span><span class="sxs-lookup"><span data-stu-id="8c13f-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-401">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-401">Compose mode</span></span>

<span data-ttu-id="8c13f-402">A propriedade `optionalAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes opcionais de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8c13f-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-403">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-403">Type:</span></span>

*   <span data-ttu-id="8c13f-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-405">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-405">Requirements</span></span>

|<span data-ttu-id="8c13f-406">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-406">Requirement</span></span>| <span data-ttu-id="8c13f-407">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-408">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-409">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-409">1.0</span></span>|
|[<span data-ttu-id="8c13f-410">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-411">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-412">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-413">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-414">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="8c13f-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c13f-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c13f-p124">Obtém o endereço de email do organizador da reunião para uma reunião especificada. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-418">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-418">Type:</span></span>

*   [<span data-ttu-id="8c13f-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c13f-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c13f-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-420">Requirements</span></span>

|<span data-ttu-id="8c13f-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-421">Requirement</span></span>| <span data-ttu-id="8c13f-422">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-423">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-424">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-424">1.0</span></span>|
|[<span data-ttu-id="8c13f-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-426">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-427">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-428">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="8c13f-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="8c13f-431">Fornece acesso aos participantes obrigatórios de um evento.</span><span class="sxs-lookup"><span data-stu-id="8c13f-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8c13f-432">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-433">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-433">Read mode</span></span>

<span data-ttu-id="8c13f-434">A propriedade `requiredAttendees` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada participante obrigatório da reunião.</span><span class="sxs-lookup"><span data-stu-id="8c13f-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-435">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-435">Compose mode</span></span>

<span data-ttu-id="8c13f-436">A propriedade `requiredAttendees` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os participantes obrigatórios de uma reunião.</span><span class="sxs-lookup"><span data-stu-id="8c13f-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-437">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-437">Type:</span></span>

*   <span data-ttu-id="8c13f-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-439">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-439">Requirements</span></span>

|<span data-ttu-id="8c13f-440">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-440">Requirement</span></span>| <span data-ttu-id="8c13f-441">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-442">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-443">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-443">1.0</span></span>|
|[<span data-ttu-id="8c13f-444">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-445">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-446">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-447">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-448">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="8c13f-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8c13f-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="8c13f-p126">Obtém o endereço de email do remetente de uma mensagem de email. Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8c13f-p127">As propriedades [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) e `sender` representam a mesma pessoa, a menos que a mensagem seja enviada por um delegado. Nesse caso, a propriedade `from` representa o delegante, e a propriedade sender, o delegado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-454">A propriedade `recipientType` do objeto `EmailAddressDetails` na propriedade `sender` é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-455">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-455">Type:</span></span>

*   [<span data-ttu-id="8c13f-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8c13f-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8c13f-457">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-457">Requirements</span></span>

|<span data-ttu-id="8c13f-458">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-458">Requirement</span></span>| <span data-ttu-id="8c13f-459">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-460">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-461">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-461">1.0</span></span>|
|[<span data-ttu-id="8c13f-462">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-463">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-464">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-465">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-466">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="8c13f-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c13f-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="8c13f-468">Obtém ou define a data e a hora em que o compromisso deve começar.</span><span class="sxs-lookup"><span data-stu-id="8c13f-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8c13f-p128">A propriedade `start` é expressa como um valor de data e hora no Tempo Universal Coordenado (UTC). Você pode usar o método [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) para converter o valor para a data e a hora local do cliente.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-471">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-471">Read mode</span></span>

<span data-ttu-id="8c13f-472">A propriedade `start` retorna um objeto `Date`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-473">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-473">Compose mode</span></span>

<span data-ttu-id="8c13f-474">A propriedade `start` retorna um objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8c13f-475">Ao usar o método [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) para definir a hora de início, deve-se usar o método [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) para converter a hora local no cliente para UTC para o servidor.</span><span class="sxs-lookup"><span data-stu-id="8c13f-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-476">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-476">Type:</span></span>

*   <span data-ttu-id="8c13f-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="8c13f-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-478">Requirements</span></span>

|<span data-ttu-id="8c13f-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-479">Requirement</span></span>| <span data-ttu-id="8c13f-480">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-481">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-482">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-482">1.0</span></span>|
|[<span data-ttu-id="8c13f-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-484">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-485">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-486">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-487">Example</span></span>

<span data-ttu-id="8c13f-488">O exemplo a seguir define a hora de início de um compromisso no modo de redação usando o método [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) do objeto `Time`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="8c13f-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8c13f-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="8c13f-490">Obtém ou define a descrição que aparece no campo de assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8c13f-491">A propriedade `subject` obtém ou define o assunto completo do item, conforme enviado pelo servidor de email.</span><span class="sxs-lookup"><span data-stu-id="8c13f-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-492">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-492">Read mode</span></span>

<span data-ttu-id="8c13f-p129">A propriedade `subject` retorna uma sequência de caracteres (string). Use a propriedade [`normalizedSubject`](#normalizedsubject-string) para obter o assunto sem os prefixos iniciais, como `RE:` e `FW:`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-495">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-495">Compose mode</span></span>

<span data-ttu-id="8c13f-496">A propriedade `subject` retorna um objeto `Subject` que fornece métodos para obter e definir o assunto.</span><span class="sxs-lookup"><span data-stu-id="8c13f-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8c13f-497">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-497">Type:</span></span>

*   <span data-ttu-id="8c13f-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8c13f-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-499">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-499">Requirements</span></span>

|<span data-ttu-id="8c13f-500">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-500">Requirement</span></span>| <span data-ttu-id="8c13f-501">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-502">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-503">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-503">1.0</span></span>|
|[<span data-ttu-id="8c13f-504">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-505">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-506">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-507">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="8c13f-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="8c13f-509">Fornece acesso aos destinatários na linha **Para** de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8c13f-510">O tipo de objeto e o nível de acesso dependem do modo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8c13f-511">Modo de leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-511">Read mode</span></span>

<span data-ttu-id="8c13f-p131">A propriedade `to` retorna uma matriz que contém um objeto `EmailAddressDetails` para cada destinatário listado na linha **Para** da mensagem. A coleção é limitada a um número máximo de 100 membros.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8c13f-514">Modo de redação</span><span class="sxs-lookup"><span data-stu-id="8c13f-514">Compose mode</span></span>

<span data-ttu-id="8c13f-515">A propriedade `to` retorna um objeto `Recipients` que fornece métodos para obter ou atualizar os destinatários na linha  **Para** da mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8c13f-516">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8c13f-516">Type:</span></span>

*   <span data-ttu-id="8c13f-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8c13f-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-518">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-518">Requirements</span></span>

|<span data-ttu-id="8c13f-519">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-519">Requirement</span></span>| <span data-ttu-id="8c13f-520">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-521">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-522">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-522">1.0</span></span>|
|[<span data-ttu-id="8c13f-523">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-524">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-525">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-526">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-527">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8c13f-528">Métodos</span><span class="sxs-lookup"><span data-stu-id="8c13f-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8c13f-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c13f-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8c13f-530">Adiciona um arquivo a uma mensagem ou a um compromisso em forma de anexo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8c13f-531">O método `addFileAttachmentAsync` carrega o arquivo no URI especificado e anexa-o ao item no formulário de redação.</span><span class="sxs-lookup"><span data-stu-id="8c13f-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8c13f-532">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8c13f-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-533">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-533">Parameters:</span></span>

|<span data-ttu-id="8c13f-534">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-534">Name</span></span>| <span data-ttu-id="8c13f-535">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-535">Type</span></span>| <span data-ttu-id="8c13f-536">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-536">Attributes</span></span>| <span data-ttu-id="8c13f-537">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8c13f-538">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-538">String</span></span>||<span data-ttu-id="8c13f-p132">O URI que fornece a localização do arquivo a ser anexado à mensagem ou ao compromisso. O comprimento máximo é de 2048 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8c13f-541">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-541">String</span></span>||<span data-ttu-id="8c13f-p133">O nome do anexo que é mostrado enquanto o anexo está sendo carregado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8c13f-544">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-544">Object</span></span>| <span data-ttu-id="8c13f-545">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-545">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-546">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-547">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-547">Object</span></span>| <span data-ttu-id="8c13f-548">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-548">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-549">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c13f-550">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-550">function</span></span>| <span data-ttu-id="8c13f-551">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-551">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-552">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c13f-553">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8c13f-554">Se houver falha ao carregar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c13f-555">Erros</span><span class="sxs-lookup"><span data-stu-id="8c13f-555">Errors</span></span>

| <span data-ttu-id="8c13f-556">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8c13f-556">Error code</span></span> | <span data-ttu-id="8c13f-557">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8c13f-558">O anexo é maior do que permitido.</span><span class="sxs-lookup"><span data-stu-id="8c13f-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8c13f-559">O anexo tem uma extensão que não é permitida.</span><span class="sxs-lookup"><span data-stu-id="8c13f-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8c13f-560">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8c13f-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-561">Requirements</span></span>

|<span data-ttu-id="8c13f-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-562">Requirement</span></span>| <span data-ttu-id="8c13f-563">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-564">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-565">1.1</span><span class="sxs-lookup"><span data-stu-id="8c13f-565">1.1</span></span>|
|[<span data-ttu-id="8c13f-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-568">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-569">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8c13f-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c13f-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8c13f-572">Adiciona um item do Exchange, como uma mensagem, em forma de anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8c13f-p134">O método `addItemAttachmentAsync` anexa o item com o identificador especificado  do Exchange ao item no formulário de redação. Se você especificar um método de retorno de chamada, o método será chamado com um parâmetro  `asyncResult` que contém o identificador do anexo ou um código que indica qualquer erro que tenha ocorrido ao anexar o item. Você pode usar o parâmetro `options` para passar informações de estado ao método de retorno de chamada, se necessário.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8c13f-576">Posteriormente, você poderá usar o identificador com o método [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) para remover o anexo na mesma sessão.</span><span class="sxs-lookup"><span data-stu-id="8c13f-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8c13f-577">Se o suplemento do Office estiver em execução no Outlook Web App, o método `addItemAttachmentAsync` pode anexar itens aos itens que não sejam aqueles que você esteja editando. No entanto, isso não é suportado e não é recomendado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-578">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-578">Parameters:</span></span>

|<span data-ttu-id="8c13f-579">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-579">Name</span></span>| <span data-ttu-id="8c13f-580">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-580">Type</span></span>| <span data-ttu-id="8c13f-581">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-581">Attributes</span></span>| <span data-ttu-id="8c13f-582">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8c13f-583">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-583">String</span></span>||<span data-ttu-id="8c13f-p135">O identificador do Exchange do item a ser anexado. O comprimento máximo é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8c13f-586">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-586">String</span></span>||<span data-ttu-id="8c13f-p136">O assunto do item a ser anexado. O comprimento máximo é de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8c13f-589">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-589">Object</span></span>| <span data-ttu-id="8c13f-590">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-590">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-591">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-592">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-592">Object</span></span>| <span data-ttu-id="8c13f-593">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-593">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-594">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c13f-595">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-595">function</span></span>| <span data-ttu-id="8c13f-596">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-596">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-597">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c13f-598">Em caso de êxito, o identificador do anexo será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8c13f-599">Se houver falha ao adicionar o anexo, o objeto `asyncResult` conterá um objeto `Error` que fornecerá uma descrição do erro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c13f-600">Erros</span><span class="sxs-lookup"><span data-stu-id="8c13f-600">Errors</span></span>

| <span data-ttu-id="8c13f-601">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8c13f-601">Error code</span></span> | <span data-ttu-id="8c13f-602">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8c13f-603">A mensagem ou o compromisso tem muitos anexos.</span><span class="sxs-lookup"><span data-stu-id="8c13f-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-604">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-604">Requirements</span></span>

|<span data-ttu-id="8c13f-605">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-605">Requirement</span></span>| <span data-ttu-id="8c13f-606">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-607">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-608">1.1</span><span class="sxs-lookup"><span data-stu-id="8c13f-608">1.1</span></span>|
|[<span data-ttu-id="8c13f-609">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-611">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-612">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-613">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-613">Example</span></span>

<span data-ttu-id="8c13f-614">O exemplo a seguir adiciona um item existente do Outlook como um anexo com o nome `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="8c13f-615">close()</span><span class="sxs-lookup"><span data-stu-id="8c13f-615">close()</span></span>

<span data-ttu-id="8c13f-616">Fecha o item atual que está sendo redigido.</span><span class="sxs-lookup"><span data-stu-id="8c13f-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="8c13f-p137">O comportamento do método `close` depende do estado atual do item que está sendo redigido. Se o item possuir alterações não salvas, o cliente solicitará que o usuário salve, descarte ou cancele a ação ao fechar.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-619">No Outlook na Web, se o item for um compromisso e tiver sido salvo anteriormente usando `saveAsync`, será solicitado ao usuário para salvar, descartar ou cancelar, mesmo se nenhuma alteração tenha ocorrido após o item ter sido salvo pela última vez.</span><span class="sxs-lookup"><span data-stu-id="8c13f-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="8c13f-620">No cliente do Outlook para área de trabalho, se a mensagem for uma resposta embutida, o método `close` não terá efeito.</span><span class="sxs-lookup"><span data-stu-id="8c13f-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-621">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-621">Requirements</span></span>

|<span data-ttu-id="8c13f-622">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-622">Requirement</span></span>| <span data-ttu-id="8c13f-623">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-624">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-625">1.3</span><span class="sxs-lookup"><span data-stu-id="8c13f-625">1.3</span></span>|
|[<span data-ttu-id="8c13f-626">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-627">Restrito</span><span class="sxs-lookup"><span data-stu-id="8c13f-627">Restricted</span></span>|
|[<span data-ttu-id="8c13f-628">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-629">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8c13f-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8c13f-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8c13f-631">Exibe um formulário de resposta que inclui o remetente e todos os destinatários da mensagem selecionada ou o organizador e todos os participantes do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-632">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c13f-633">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8c13f-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8c13f-634">Se qualquer um dos parâmetros da sequência de caracteres exceder o limite, `displayReplyAllForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8c13f-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8c13f-p138">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-638">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-638">Parameters:</span></span>

|<span data-ttu-id="8c13f-639">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-639">Name</span></span>| <span data-ttu-id="8c13f-640">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-640">Type</span></span>| <span data-ttu-id="8c13f-641">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8c13f-642">String | Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-642">String &#124; Object</span></span>| |<span data-ttu-id="8c13f-p139">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8c13f-645">**OU**</span><span class="sxs-lookup"><span data-stu-id="8c13f-645">**OR**</span></span><br/><span data-ttu-id="8c13f-p140">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8c13f-648">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-648">String</span></span> | <span data-ttu-id="8c13f-649">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-649">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-p141">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8c13f-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8c13f-653">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-653">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-654">Uma matriz de objetos JSON que são anexos do arquivo ou do item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8c13f-655">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-655">String</span></span> | | <span data-ttu-id="8c13f-p142">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8c13f-658">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-658">String</span></span> | | <span data-ttu-id="8c13f-659">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8c13f-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8c13f-660">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-660">String</span></span> | | <span data-ttu-id="8c13f-p143">Usado somente se `type` estiver definido como `file`. O URI da localização para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8c13f-663">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-663">String</span></span> | | <span data-ttu-id="8c13f-p144">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Essa é uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8c13f-667">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-667">function</span></span> | <span data-ttu-id="8c13f-668">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-668">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-669">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-670">Requirements</span></span>

|<span data-ttu-id="8c13f-671">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-671">Requirement</span></span>| <span data-ttu-id="8c13f-672">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-673">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-674">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-674">1.0</span></span>|
|[<span data-ttu-id="8c13f-675">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-676">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-677">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-678">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c13f-679">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8c13f-679">Examples</span></span>

<span data-ttu-id="8c13f-680">O código a seguir transmite uma sequência de caracteres para a função `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8c13f-681">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8c13f-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8c13f-682">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8c13f-683">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-683">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="8c13f-684">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-684">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="8c13f-685">Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8c13f-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8c13f-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="8c13f-687">Exibe um formulário de resposta que inclui o remetente da mensagem selecionada ou o organizador do compromisso selecionado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-688">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c13f-689">No Outlook Web App, o formulário de resposta é exibido como um formulário pop-out no modo de exibição de três colunas e um formulário pop-up no modo de exibição de uma ou duas colunas.</span><span class="sxs-lookup"><span data-stu-id="8c13f-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8c13f-690">Se qualquer um dos parâmetros da sequência de caracteres exceder o limite, `displayReplyForm` gerará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="8c13f-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8c13f-p145">Quando os anexos são especificados no parâmetro `formData.attachments`, o Outlook e o Outlook Web App tentam baixar todos os anexos e anexá-los ao formulário de resposta. Se a adição de anexos falhar, será exibido um erro na interface de usuário do formulário. Se isso não for possível, nenhuma mensagem de erro será apresentada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-694">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-694">Parameters:</span></span>

|<span data-ttu-id="8c13f-695">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-695">Name</span></span>| <span data-ttu-id="8c13f-696">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-696">Type</span></span>| <span data-ttu-id="8c13f-697">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8c13f-698">String | Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-698">String &#124; Object</span></span>| | <span data-ttu-id="8c13f-p146">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8c13f-701">**OU**</span><span class="sxs-lookup"><span data-stu-id="8c13f-701">**OR**</span></span><br/><span data-ttu-id="8c13f-p147">Um objeto que contém os dados do corpo ou do anexo e uma função de retorno de chamada. O objeto é definido da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8c13f-704">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-704">String</span></span> | <span data-ttu-id="8c13f-705">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-705">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-p148">Uma sequência de caracteres que contém texto e HTML e que representa o corpo do formulário de resposta. A sequência de caracteres está limitada a 32 KB.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8c13f-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8c13f-709">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-709">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-710">Uma matriz de objetos JSON que são anexos do arquivo ou do item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8c13f-711">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-711">String</span></span> | | <span data-ttu-id="8c13f-p149">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8c13f-714">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-714">String</span></span> | | <span data-ttu-id="8c13f-715">Uma sequência de caracteres que contém o nome do anexo, com até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="8c13f-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8c13f-716">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-716">String</span></span> | | <span data-ttu-id="8c13f-p150">Usado somente se `type` estiver definido como `file`. O URI da localização para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8c13f-719">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-719">String</span></span> | | <span data-ttu-id="8c13f-p151">Usado somente se `type` estiver definido como `item`. A ID do item do EWS do anexo. Essa é uma sequência de caracteres de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8c13f-723">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-723">function</span></span> | <span data-ttu-id="8c13f-724">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-724">&lt;optional&gt;</span></span> | <span data-ttu-id="8c13f-725">Quando o método for concluído, a função passada ao parâmetro `callback` será chamada com um único parâmetro  `asyncResult`, que é um objeto [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-726">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-726">Requirements</span></span>

|<span data-ttu-id="8c13f-727">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-727">Requirement</span></span>| <span data-ttu-id="8c13f-728">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-729">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-730">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-730">1.0</span></span>|
|[<span data-ttu-id="8c13f-731">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-732">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-733">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-734">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c13f-735">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8c13f-735">Examples</span></span>

<span data-ttu-id="8c13f-736">O código a seguir transmite uma sequência de caracteres para a função `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8c13f-737">Responder com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="8c13f-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8c13f-738">Responder apenas com um corpo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8c13f-739">Responder com um corpo e um anexo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-739">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="8c13f-740">Responder com um corpo e um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-740">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="8c13f-741">Responder com um corpo, um anexo de arquivo, um anexo de item e um retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="8c13f-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8c13f-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="8c13f-743">Obtém as entidades encontradas no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-744">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-745">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-745">Requirements</span></span>

|<span data-ttu-id="8c13f-746">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-746">Requirement</span></span>| <span data-ttu-id="8c13f-747">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-748">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-749">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-749">1.0</span></span>|
|[<span data-ttu-id="8c13f-750">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-751">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-752">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-753">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-754">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-754">Returns:</span></span>

<span data-ttu-id="8c13f-755">Tipo: [Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8c13f-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8c13f-756">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-756">Example</span></span>

<span data-ttu-id="8c13f-757">O exemplo a seguir acessa as entidades de contatos no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="8c13f-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8c13f-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8c13f-759">Obtém uma matriz de todas as entidades do tipo de entidade especificado encontrado no corpo do item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-760">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-761">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-761">Parameters:</span></span>

|<span data-ttu-id="8c13f-762">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-762">Name</span></span>| <span data-ttu-id="8c13f-763">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-763">Type</span></span>| <span data-ttu-id="8c13f-764">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8c13f-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8c13f-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="8c13f-766">Um dos valores de enumeração de EntityType.</span><span class="sxs-lookup"><span data-stu-id="8c13f-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-767">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-767">Requirements</span></span>

|<span data-ttu-id="8c13f-768">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-768">Requirement</span></span>| <span data-ttu-id="8c13f-769">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-770">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-771">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-771">1.0</span></span>|
|[<span data-ttu-id="8c13f-772">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-773">Restrito</span><span class="sxs-lookup"><span data-stu-id="8c13f-773">Restricted</span></span>|
|[<span data-ttu-id="8c13f-774">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-775">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-776">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-776">Returns:</span></span>

<span data-ttu-id="8c13f-777">Se o valor passado em `entityType` não for um membro válido da enumeração `EntityType`, o método retornará nulo.</span><span class="sxs-lookup"><span data-stu-id="8c13f-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8c13f-778">Se nenhuma entidade do tipo especificado estiver presente no corpo do item, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8c13f-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="8c13f-779">Caso contrário, o tipo dos objetos na matriz retornada dependerá do tipo de entidade solicitado no parâmetro `entityType`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8c13f-780">Embora o nível de permissão mínimo para usar esse método seja **Restricted**, alguns tipos de entidade exigem a permissão **ReadItem** para obter acesso, conforme especificado na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8c13f-781">Valor de `entityType`</span><span class="sxs-lookup"><span data-stu-id="8c13f-781">Value of `entityType`</span></span> | <span data-ttu-id="8c13f-782">Tipo de objetos na matriz retornada</span><span class="sxs-lookup"><span data-stu-id="8c13f-782">Type of objects in returned array</span></span> | <span data-ttu-id="8c13f-783">Nível de Permissão Exigido</span><span class="sxs-lookup"><span data-stu-id="8c13f-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8c13f-784">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-784">String</span></span> | <span data-ttu-id="8c13f-785">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8c13f-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8c13f-786">Contact</span><span class="sxs-lookup"><span data-stu-id="8c13f-786">Contact</span></span> | <span data-ttu-id="8c13f-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c13f-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8c13f-788">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-788">String</span></span> | <span data-ttu-id="8c13f-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c13f-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8c13f-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8c13f-790">MeetingSuggestion</span></span> | <span data-ttu-id="8c13f-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c13f-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8c13f-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8c13f-792">PhoneNumber</span></span> | <span data-ttu-id="8c13f-793">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8c13f-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8c13f-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8c13f-794">TaskSuggestion</span></span> | <span data-ttu-id="8c13f-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8c13f-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8c13f-796">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-796">String</span></span> | <span data-ttu-id="8c13f-797">**Restrito**</span><span class="sxs-lookup"><span data-stu-id="8c13f-797">**Restricted**</span></span> |

<span data-ttu-id="8c13f-798">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8c13f-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8c13f-799">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-799">Example</span></span>

<span data-ttu-id="8c13f-800">O exemplo a seguir mostra como acessar uma matriz de sequências de caracteres que representa endereços postais no corpo do item atual.</span><span class="sxs-lookup"><span data-stu-id="8c13f-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="8c13f-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8c13f-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8c13f-802">Retorna entidades bem conhecidas no item selecionado que passam o filtro nomeado definido no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-803">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c13f-804">O método `getFilteredEntitiesByName` retorna as entidades que correspondem à expressão regular definida no elemento de regra [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) no arquivo de manifesto XML com o valor do elemento `FilterName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-805">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-805">Parameters:</span></span>

|<span data-ttu-id="8c13f-806">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-806">Name</span></span>| <span data-ttu-id="8c13f-807">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-807">Type</span></span>| <span data-ttu-id="8c13f-808">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8c13f-809">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-809">String</span></span>|<span data-ttu-id="8c13f-810">O nome do elemento de regra `ItemHasKnownEntity` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="8c13f-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-811">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-811">Requirements</span></span>

|<span data-ttu-id="8c13f-812">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-812">Requirement</span></span>| <span data-ttu-id="8c13f-813">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-814">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-815">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-815">1.0</span></span>|
|[<span data-ttu-id="8c13f-816">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-817">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-818">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-819">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-820">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-820">Returns:</span></span>

<span data-ttu-id="8c13f-p153">Se não houver nenhum elemento `ItemHasKnownEntity` no manifesto com um valor de elemento `FilterName` que corresponda ao parâmetro `name`, o método retornará `null`. Se o parâmetro `name` corresponder a um elemento `ItemHasKnownEntity` no manifesto, mas não houver entidades no item atual que correspondam, o método retornará uma matriz vazia.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8c13f-823">Tipo: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8c13f-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="8c13f-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8c13f-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8c13f-825">Retorna valores de sequência de caracteres no item selecionado que correspondem às expressões regulares definidas no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-826">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c13f-p154">O método `getRegExMatches` retorna sequências de caracteres que correspondem à expressão regular definida em cada elemento de regra  `ItemHasRegularExpressionMatch` ou  `ItemHasKnownEntity` no arquivo de manifesto XML. Para uma regra `ItemHasRegularExpressionMatch`, uma sequência de caracteres correspondente deve ocorrer na propriedade do item especificada por essa regra. O tipo simples `PropertyName` define as propriedades compatíveis.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8c13f-830">Por exemplo, considere que um manifesto de suplemento tenha o seguinte elemento `Rule` :</span><span class="sxs-lookup"><span data-stu-id="8c13f-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8c13f-831">O objeto retornado por `getRegExMatches` teria duas propriedades: `fruits` e `veggies`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="8c13f-p155">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar ainda mais o corpo e não tentará retornar o corpo inteiro do item. Usar uma expressão regular como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados. Em vez disso, use o método [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) para recuperar o corpo inteiro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8c13f-835">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-835">Requirements</span></span>

|<span data-ttu-id="8c13f-836">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-836">Requirement</span></span>| <span data-ttu-id="8c13f-837">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-838">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-839">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-839">1.0</span></span>|
|[<span data-ttu-id="8c13f-840">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-841">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-842">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-843">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-844">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-844">Returns:</span></span>

<span data-ttu-id="8c13f-p156">Um objeto que contém matrizes de sequências de caracteres que correspondem às expressões regulares definidas no arquivo de manifesto XML. O nome de cada matriz é igual ao valor correspondente do atributo `RegExName` da regra `ItemHasRegularExpressionMatch` correspondente ou do atributo `FilterName` da regra `ItemHasKnownEntity` correspondente.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8c13f-847">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c13f-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c13f-848">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c13f-849">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-849">Example</span></span>

<span data-ttu-id="8c13f-850">O exemplo a seguir mostra como acessar a matriz de correspondências para os <rule>elementos da expressão regular `fruits` e `veggies` que são especificados no manifesto.</rule></span><span class="sxs-lookup"><span data-stu-id="8c13f-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8c13f-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8c13f-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8c13f-852">Retorna valores de sequência de caracteres no item selecionado que correspondem à expressão regular nomeada definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-853">Esse método não é compatível com o Outlook para iOS ou o Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="8c13f-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8c13f-854">O método `getRegExMatchesByName` retorna as sequências de caracteres que correspondem à expressão regular definida no elemento de regra `ItemHasRegularExpressionMatch` no arquivo de manifesto XML com o valor de elemento `RegExName` especificado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8c13f-p157">Se você especificar uma regra `ItemHasRegularExpressionMatch` na propriedade do corpo de um item, a expressão regular deverá filtrar ainda mais o corpo e não tentará retornar o corpo inteiro do item. Usar uma expressão regular  como `.*` para obter o corpo inteiro de um item nem sempre retornará os resultados esperados.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-857">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-857">Parameters:</span></span>

|<span data-ttu-id="8c13f-858">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-858">Name</span></span>| <span data-ttu-id="8c13f-859">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-859">Type</span></span>| <span data-ttu-id="8c13f-860">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8c13f-861">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-861">String</span></span>|<span data-ttu-id="8c13f-862">O nome do elemento de regra `ItemHasRegularExpressionMatch` que define o filtro a ser correspondido.</span><span class="sxs-lookup"><span data-stu-id="8c13f-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-863">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-863">Requirements</span></span>

|<span data-ttu-id="8c13f-864">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-864">Requirement</span></span>| <span data-ttu-id="8c13f-865">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-866">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-867">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-867">1.0</span></span>|
|[<span data-ttu-id="8c13f-868">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-869">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-870">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-871">Leitura</span><span class="sxs-lookup"><span data-stu-id="8c13f-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-872">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-872">Returns:</span></span>

<span data-ttu-id="8c13f-873">Uma matriz que contém as sequências de caracteres que correspondem à expressão regular definida no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8c13f-874">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c13f-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c13f-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8c13f-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c13f-876">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8c13f-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8c13f-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8c13f-878">Retorna de forma assíncrona os dados selecionados do assunto ou do corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8c13f-p158">Se não houver seleção, mas o cursor estiver no corpo ou no assunto, o método retornará nulo para os dados selecionados. Se um campo que não seja o corpo ou o assunto estiver selecionado, o método retornará o erro `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-881">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-881">Parameters:</span></span>

|<span data-ttu-id="8c13f-882">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-882">Name</span></span>| <span data-ttu-id="8c13f-883">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-883">Type</span></span>| <span data-ttu-id="8c13f-884">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-884">Attributes</span></span>| <span data-ttu-id="8c13f-885">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8c13f-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8c13f-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8c13f-p159">Solicita um formato para os dados. Se for Text, o método retornará o texto sem formatação em forma de sequência de caracteres, removendo quaisquer tags HTML presentes. Se for HTML, o método retornará o texto selecionado, seja ele texto sem formatação ou HTML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8c13f-890">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-890">Object</span></span>| <span data-ttu-id="8c13f-891">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-891">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-892">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-893">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-893">Object</span></span>| <span data-ttu-id="8c13f-894">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-894">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-895">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c13f-896">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-896">function</span></span>||<span data-ttu-id="8c13f-897">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c13f-898">Para acessar os dados selecionados do método de retorno de chamada, chame `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8c13f-899">Para acessar a propriedade original de onde a seleção é proveniente, chame `asyncResult.value.sourceProperty`, que será `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-900">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-900">Requirements</span></span>

|<span data-ttu-id="8c13f-901">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-901">Requirement</span></span>| <span data-ttu-id="8c13f-902">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-903">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-904">1.2</span><span class="sxs-lookup"><span data-stu-id="8c13f-904">1.2</span></span>|
|[<span data-ttu-id="8c13f-905">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-907">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-908">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8c13f-909">Retorna:</span><span class="sxs-lookup"><span data-stu-id="8c13f-909">Returns:</span></span>

<span data-ttu-id="8c13f-910">Os dados selecionados em forma de sequência de caracteres com formato determinado por `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8c13f-911">

<dt>Tipo</dt>

</span><span class="sxs-lookup"><span data-stu-id="8c13f-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8c13f-912">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8c13f-913">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-913">Example</span></span>

```
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8c13f-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8c13f-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8c13f-915">Carrega de forma assíncrona as propriedades personalizadas para esse suplemento no item selecionado.</span><span class="sxs-lookup"><span data-stu-id="8c13f-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8c13f-p161">Propriedades personalizadas são armazenadas como pares chave/valor de acordo com o aplicativo e o item. Este método retornará um objeto `CustomProperties` no retorno de chamada, que fornece métodos para acessar as propriedades personalizadas específicas para o item atual e o suplemento atual. As propriedades personalizadas não são criptografadas no item, portanto, isto não deve ser usado como armazenamento seguro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-919">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-919">Parameters:</span></span>

|<span data-ttu-id="8c13f-920">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-920">Name</span></span>| <span data-ttu-id="8c13f-921">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-921">Type</span></span>| <span data-ttu-id="8c13f-922">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-922">Attributes</span></span>| <span data-ttu-id="8c13f-923">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8c13f-924">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-924">function</span></span>||<span data-ttu-id="8c13f-925">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c13f-926">As propriedades personalizadas são fornecidas como um objeto [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8c13f-927">Esse objeto pode ser usado para obter, definir e remover propriedades personalizadas do item e salvar as alterações no conjunto de propriedades personalizadas de volta no servidor.</span><span class="sxs-lookup"><span data-stu-id="8c13f-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8c13f-928">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-928">Object</span></span>| <span data-ttu-id="8c13f-929">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-929">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-930">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="8c13f-931">Esse objeto pode ser acessado pela propriedade `asyncResult.asyncContext` na função de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-932">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-932">Requirements</span></span>

|<span data-ttu-id="8c13f-933">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-933">Requirement</span></span>| <span data-ttu-id="8c13f-934">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-935">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-936">1.0</span><span class="sxs-lookup"><span data-stu-id="8c13f-936">1.0</span></span>|
|[<span data-ttu-id="8c13f-937">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-938">ReadItem</span></span>|
|[<span data-ttu-id="8c13f-939">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-940">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="8c13f-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-941">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-941">Example</span></span>

<span data-ttu-id="8c13f-p164">O exemplo de código a seguir mostra como usar o método `loadCustomPropertiesAsync` para carregar de forma assíncrona as propriedades personalizadas que são específicas para o item atual. O exemplo também mostra como usar o método `CustomProperties.saveAsync` para salvar essas propriedades de volta ao servidor. Depois de carregar as propriedades personalizadas, o exemplo de código usará o método `CustomProperties.get` para ler a propriedade personalizada `myProp`, o método `CustomProperties.set` para gravar a propriedade personalizada `otherProp` e chamará o método `saveAsync` para salvar as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8c13f-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8c13f-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8c13f-946">Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c13f-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8c13f-p165">O método `removeAttachmentAsync` remove o anexo com o identificador especificado do item. Como prática recomendada, deve-se usar o identificador do anexo para remover um anexo somente se o mesmo aplicativo de email tiver adicionado esse anexo na mesma sessão. No Outlook Web App e no OWA para Dispositivos, o identificador do anexo é válido apenas dentro da mesma sessão. Uma sessão é finalizada quando o usuário fecha o aplicativo ou quando o usuário começa a redigir em um formulário embutido e, subsequentemente, sai do formulário embutido para continuar em uma janela separada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-951">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-951">Parameters:</span></span>

|<span data-ttu-id="8c13f-952">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-952">Name</span></span>| <span data-ttu-id="8c13f-953">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-953">Type</span></span>| <span data-ttu-id="8c13f-954">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-954">Attributes</span></span>| <span data-ttu-id="8c13f-955">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8c13f-956">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-956">String</span></span>||<span data-ttu-id="8c13f-p166">O identificador do anexo a ser removido. O comprimento máximo da sequência de caracteres é de 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="8c13f-959">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-959">Object</span></span>| <span data-ttu-id="8c13f-960">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-960">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-961">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-962">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-962">Object</span></span>| <span data-ttu-id="8c13f-963">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-963">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-964">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c13f-965">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-965">function</span></span>| <span data-ttu-id="8c13f-966">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-966">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-967">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8c13f-968">Se a remoção do anexo falhar, a propriedade `asyncResult.error` conterá um código de erro com o motivo da falha.</span><span class="sxs-lookup"><span data-stu-id="8c13f-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8c13f-969">Erros</span><span class="sxs-lookup"><span data-stu-id="8c13f-969">Errors</span></span>

| <span data-ttu-id="8c13f-970">Código de erro</span><span class="sxs-lookup"><span data-stu-id="8c13f-970">Error code</span></span> | <span data-ttu-id="8c13f-971">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8c13f-972">O identificador do anexo não existe.</span><span class="sxs-lookup"><span data-stu-id="8c13f-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-973">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-973">Requirements</span></span>

|<span data-ttu-id="8c13f-974">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-974">Requirement</span></span>| <span data-ttu-id="8c13f-975">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-976">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-977">1.1</span><span class="sxs-lookup"><span data-stu-id="8c13f-977">1.1</span></span>|
|[<span data-ttu-id="8c13f-978">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-980">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-981">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-982">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-982">Example</span></span>

<span data-ttu-id="8c13f-983">O código a seguir remove um anexo com um identificador '0'.</span><span class="sxs-lookup"><span data-stu-id="8c13f-983">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="8c13f-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="8c13f-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="8c13f-985">Salva um item de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="8c13f-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="8c13f-p167">Quando chamado, este método salva a mensagem atual como um rascunho e retorna o identificador do item por meio do método de retorno de chamada. No Outlook Web App ou no Outlook no modo online, o item é salvo no servidor. No Outlook no modo cache, o item é salvo no cache local.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-989">Se o seu suplemento chamar `saveAsync` em um item no modo de redação para obter um `itemId` para usar com o EWS ou a API REST, esteja ciente de que quando o Outlook estiver no modo de cache, poderá levar algum tempo antes do item realmente ser sincronizado com o servidor.</span><span class="sxs-lookup"><span data-stu-id="8c13f-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="8c13f-990">Até que o item seja sincronizado, o uso de `itemId` retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="8c13f-p169">Como compromissos não têm um estado de rascunho, se `saveAsync` for chamado em um compromisso no modo de redação, o item será salvo como um compromisso normal no calendário do usuário. Para novos compromissos que não foram salvos antes, nenhum convite será enviado. Salvar um compromisso existente enviará uma atualização aos participantes adicionados ou removidos.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="8c13f-994">Os seguintes clientes possuem um comportamento diferente para o `saveAsync` nos compromissos no modo de redação:</span><span class="sxs-lookup"><span data-stu-id="8c13f-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="8c13f-995">Mac Outlook não suporta `saveAsync` em uma reunião no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8c13f-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="8c13f-996">Chamar `saveAsync` em uma reunião no Outlook Mac retornará um erro.</span><span class="sxs-lookup"><span data-stu-id="8c13f-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="8c13f-997">O Outlook na Web sempre enviará um convite ou atualização quando `saveAsync` for chamado em um compromisso no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="8c13f-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-998">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-998">Parameters:</span></span>

|<span data-ttu-id="8c13f-999">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-999">Name</span></span>| <span data-ttu-id="8c13f-1000">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-1000">Type</span></span>| <span data-ttu-id="8c13f-1001">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-1001">Attributes</span></span>| <span data-ttu-id="8c13f-1002">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="8c13f-1003">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-1003">Object</span></span>| <span data-ttu-id="8c13f-1004">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-1005">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-1006">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-1006">Object</span></span>| <span data-ttu-id="8c13f-1007">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-1008">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8c13f-1009">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-1009">function</span></span>||<span data-ttu-id="8c13f-1010">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8c13f-1011">Em caso de sucesso, o identificador do item será fornecido na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8c13f-1012">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-1012">Requirements</span></span>

|<span data-ttu-id="8c13f-1013">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-1013">Requirement</span></span>| <span data-ttu-id="8c13f-1014">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-1015">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="8c13f-1016">1.3</span></span>|
|[<span data-ttu-id="8c13f-1017">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-1019">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-1020">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="8c13f-1021">Exemplos</span><span class="sxs-lookup"><span data-stu-id="8c13f-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="8c13f-p171">A seguir apresentamos um exemplo do parâmetro `result` passado para a função de retorno de chamada. A propriedade `value` contém o ID do item.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8c13f-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8c13f-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8c13f-1025">Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8c13f-p172">O método `setSelectedDataAsync` insere a sequência de caracteres especificada no local do cursor no corpo ou no assunto do item ou, se o texto estiver selecionado no editor, substitui o texto selecionado. Se o cursor não estiver no campo do corpo ou do assunto, um erro será retornado. Após a inserção, o cursor será colocado no final do conteúdo inserido.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8c13f-1029">Parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8c13f-1029">Parameters:</span></span>

|<span data-ttu-id="8c13f-1030">Nome</span><span class="sxs-lookup"><span data-stu-id="8c13f-1030">Name</span></span>| <span data-ttu-id="8c13f-1031">Tipo</span><span class="sxs-lookup"><span data-stu-id="8c13f-1031">Type</span></span>| <span data-ttu-id="8c13f-1032">Atributos</span><span class="sxs-lookup"><span data-stu-id="8c13f-1032">Attributes</span></span>| <span data-ttu-id="8c13f-1033">Descrição</span><span class="sxs-lookup"><span data-stu-id="8c13f-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8c13f-1034">String</span><span class="sxs-lookup"><span data-stu-id="8c13f-1034">String</span></span>||<span data-ttu-id="8c13f-p173">Os dados a serem inseridos. Os dados não devem exceder 1.000.000 de caracteres. Se forem passados mais de 1.000.000 caracteres, ocorrerá uma exceção `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8c13f-1038">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-1038">Object</span></span>| <span data-ttu-id="8c13f-1039">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-1040">Um literal de objeto que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8c13f-1041">Object</span><span class="sxs-lookup"><span data-stu-id="8c13f-1041">Object</span></span>| <span data-ttu-id="8c13f-1042">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-1043">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="8c13f-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8c13f-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="8c13f-1045">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="8c13f-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="8c13f-p174">Se `text`, o estilo atual é aplicado no Outlook Web App e no Outlook. Se o campo for um editor de HTML, apenas os dados de texto são inseridos, mesmo se os dados forem HTML.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8c13f-p175">Se `html` e o campo forem compatíveis com HTML (e o assunto não), o estilo atual é aplicado no Outlook Web App e o estilo padrão será aplicado no Outlook. Se o campo for um campo de texto, retornará um erro `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="8c13f-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8c13f-1050">Se `coercionType` não estiver definido, o resultado depende do campo: se o campo for HTML, será usado HTML; se o campo for texto, será usado texto sem formatação.</span><span class="sxs-lookup"><span data-stu-id="8c13f-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8c13f-1051">function</span><span class="sxs-lookup"><span data-stu-id="8c13f-1051">function</span></span>||<span data-ttu-id="8c13f-1052">Quando o método for concluído, a função passada no parâmetro `callback` é chamada com um parâmetro único, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8c13f-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8c13f-1053">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8c13f-1053">Requirements</span></span>

|<span data-ttu-id="8c13f-1054">Requisito</span><span class="sxs-lookup"><span data-stu-id="8c13f-1054">Requirement</span></span>| <span data-ttu-id="8c13f-1055">Valor</span><span class="sxs-lookup"><span data-stu-id="8c13f-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="8c13f-1056">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8c13f-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8c13f-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="8c13f-1057">1.2</span></span>|
|[<span data-ttu-id="8c13f-1058">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8c13f-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8c13f-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8c13f-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="8c13f-1060">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c13f-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8c13f-1061">Redigir</span><span class="sxs-lookup"><span data-stu-id="8c13f-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8c13f-1062">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8c13f-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```