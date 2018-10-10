
# <a name="userprofile"></a><span data-ttu-id="18f33-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="18f33-101">userProfile</span></span>

### <span data-ttu-id="18f33-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="18f33-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f33-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18f33-104">Requirements</span></span>

|<span data-ttu-id="18f33-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="18f33-105">Requirement</span></span>| <span data-ttu-id="18f33-106">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f33-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18f33-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f33-108">1.0</span><span class="sxs-lookup"><span data-stu-id="18f33-108">1.0</span></span>|
|[<span data-ttu-id="18f33-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18f33-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18f33-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18f33-110">ReadItem</span></span>|
|[<span data-ttu-id="18f33-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18f33-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f33-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="18f33-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="18f33-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="18f33-113">Members and methods</span></span>

| <span data-ttu-id="18f33-114">Membro</span><span class="sxs-lookup"><span data-stu-id="18f33-114">Member</span></span> | <span data-ttu-id="18f33-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="18f33-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="18f33-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="18f33-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="18f33-117">Membro</span><span class="sxs-lookup"><span data-stu-id="18f33-117">Member</span></span> |
| [<span data-ttu-id="18f33-118">displayName</span><span class="sxs-lookup"><span data-stu-id="18f33-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="18f33-119">Membro</span><span class="sxs-lookup"><span data-stu-id="18f33-119">Member</span></span> |
| [<span data-ttu-id="18f33-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="18f33-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="18f33-121">Membro</span><span class="sxs-lookup"><span data-stu-id="18f33-121">Member</span></span> |
| [<span data-ttu-id="18f33-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="18f33-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="18f33-123">Membro</span><span class="sxs-lookup"><span data-stu-id="18f33-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="18f33-124">Membros</span><span class="sxs-lookup"><span data-stu-id="18f33-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="18f33-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="18f33-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="18f33-126">O suporte é oferecido a esse membro somente no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="18f33-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="18f33-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="18f33-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="18f33-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="18f33-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="18f33-129">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-129">Value</span></span> | <span data-ttu-id="18f33-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="18f33-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="18f33-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="18f33-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="18f33-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="18f33-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="18f33-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="18f33-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="18f33-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="18f33-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="18f33-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18f33-135">Type:</span></span>

*   <span data-ttu-id="18f33-136">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="18f33-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f33-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18f33-137">Requirements</span></span>

|<span data-ttu-id="18f33-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="18f33-138">Requirement</span></span>| <span data-ttu-id="18f33-139">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f33-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18f33-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f33-141">1.6</span><span class="sxs-lookup"><span data-stu-id="18f33-141">-16</span></span> |
|[<span data-ttu-id="18f33-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18f33-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18f33-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18f33-143">ReadItem</span></span>|
|[<span data-ttu-id="18f33-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18f33-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f33-145">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="18f33-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18f33-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18f33-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="18f33-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="18f33-147">displayName :String</span></span>

<span data-ttu-id="18f33-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="18f33-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="18f33-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18f33-149">Type:</span></span>

*   <span data-ttu-id="18f33-150">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="18f33-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f33-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18f33-151">Requirements</span></span>

|<span data-ttu-id="18f33-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="18f33-152">Requirement</span></span>| <span data-ttu-id="18f33-153">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f33-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18f33-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f33-155">1.0</span><span class="sxs-lookup"><span data-stu-id="18f33-155">1.0</span></span>|
|[<span data-ttu-id="18f33-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18f33-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18f33-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18f33-157">ReadItem</span></span>|
|[<span data-ttu-id="18f33-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18f33-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f33-159">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="18f33-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18f33-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18f33-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="18f33-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="18f33-161">emailAddress :String</span></span>

<span data-ttu-id="18f33-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="18f33-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="18f33-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18f33-163">Type:</span></span>

*   <span data-ttu-id="18f33-164">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="18f33-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f33-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18f33-165">Requirements</span></span>

|<span data-ttu-id="18f33-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="18f33-166">Requirement</span></span>| <span data-ttu-id="18f33-167">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f33-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18f33-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f33-169">1.0</span><span class="sxs-lookup"><span data-stu-id="18f33-169">1.0</span></span>|
|[<span data-ttu-id="18f33-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18f33-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18f33-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18f33-171">ReadItem</span></span>|
|[<span data-ttu-id="18f33-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18f33-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f33-173">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="18f33-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18f33-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18f33-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="18f33-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="18f33-175">timeZone :String</span></span>

<span data-ttu-id="18f33-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="18f33-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="18f33-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="18f33-177">Type:</span></span>

*   <span data-ttu-id="18f33-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="18f33-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18f33-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="18f33-179">Requirements</span></span>

|<span data-ttu-id="18f33-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="18f33-180">Requirement</span></span>| <span data-ttu-id="18f33-181">Valor</span><span class="sxs-lookup"><span data-stu-id="18f33-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="18f33-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="18f33-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18f33-183">1.0</span><span class="sxs-lookup"><span data-stu-id="18f33-183">1.0</span></span>|
|[<span data-ttu-id="18f33-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="18f33-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18f33-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18f33-185">ReadItem</span></span>|
|[<span data-ttu-id="18f33-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="18f33-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18f33-187">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="18f33-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18f33-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="18f33-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```