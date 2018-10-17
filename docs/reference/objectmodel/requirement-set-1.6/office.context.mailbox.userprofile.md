
# <a name="userprofile"></a><span data-ttu-id="484a2-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="484a2-101">userProfile</span></span>

### <span data-ttu-id="484a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="484a2-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="484a2-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="484a2-104">Requirements</span></span>

|<span data-ttu-id="484a2-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="484a2-105">Requirement</span></span>| <span data-ttu-id="484a2-106">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="484a2-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="484a2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484a2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="484a2-108">1.0</span></span>|
|[<span data-ttu-id="484a2-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="484a2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484a2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484a2-110">ReadItem</span></span>|
|[<span data-ttu-id="484a2-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="484a2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484a2-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="484a2-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="484a2-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="484a2-113">Members and methods</span></span>

| <span data-ttu-id="484a2-114">Membro</span><span class="sxs-lookup"><span data-stu-id="484a2-114">Member</span></span> | <span data-ttu-id="484a2-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="484a2-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="484a2-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="484a2-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="484a2-117">Membro</span><span class="sxs-lookup"><span data-stu-id="484a2-117">Member</span></span> |
| [<span data-ttu-id="484a2-118">displayName</span><span class="sxs-lookup"><span data-stu-id="484a2-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="484a2-119">Membro</span><span class="sxs-lookup"><span data-stu-id="484a2-119">Member</span></span> |
| [<span data-ttu-id="484a2-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="484a2-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="484a2-121">Membro</span><span class="sxs-lookup"><span data-stu-id="484a2-121">Member</span></span> |
| [<span data-ttu-id="484a2-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="484a2-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="484a2-123">Membro</span><span class="sxs-lookup"><span data-stu-id="484a2-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="484a2-124">Membros</span><span class="sxs-lookup"><span data-stu-id="484a2-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="484a2-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="484a2-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="484a2-126">Esse membro é compatível somente no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="484a2-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="484a2-p102">Obtém o tipo de conta do usuário associado com a caixa de correio. Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="484a2-p102">Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.</span></span>

| <span data-ttu-id="484a2-129">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-129">Value</span></span> | <span data-ttu-id="484a2-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="484a2-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="484a2-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="484a2-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="484a2-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="484a2-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="484a2-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="484a2-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="484a2-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="484a2-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="484a2-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="484a2-135">Type:</span></span>

*   <span data-ttu-id="484a2-136">String</span><span class="sxs-lookup"><span data-stu-id="484a2-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484a2-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="484a2-137">Requirements</span></span>

|<span data-ttu-id="484a2-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="484a2-138">Requirement</span></span>| <span data-ttu-id="484a2-139">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="484a2-140">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="484a2-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484a2-141">1.6</span><span class="sxs-lookup"><span data-stu-id="484a2-141">-16</span></span> |
|[<span data-ttu-id="484a2-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="484a2-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484a2-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484a2-143">ReadItem</span></span>|
|[<span data-ttu-id="484a2-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="484a2-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484a2-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="484a2-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484a2-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="484a2-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="484a2-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="484a2-147">displayName :String</span></span>

<span data-ttu-id="484a2-148">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="484a2-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="484a2-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="484a2-149">Type:</span></span>

*   <span data-ttu-id="484a2-150">String</span><span class="sxs-lookup"><span data-stu-id="484a2-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484a2-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="484a2-151">Requirements</span></span>

|<span data-ttu-id="484a2-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="484a2-152">Requirement</span></span>| <span data-ttu-id="484a2-153">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="484a2-154">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="484a2-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484a2-155">1.0</span><span class="sxs-lookup"><span data-stu-id="484a2-155">1.0</span></span>|
|[<span data-ttu-id="484a2-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="484a2-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484a2-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484a2-157">ReadItem</span></span>|
|[<span data-ttu-id="484a2-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="484a2-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484a2-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="484a2-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484a2-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="484a2-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="484a2-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="484a2-161">emailAddress :String</span></span>

<span data-ttu-id="484a2-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="484a2-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="484a2-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="484a2-163">Type:</span></span>

*   <span data-ttu-id="484a2-164">String</span><span class="sxs-lookup"><span data-stu-id="484a2-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484a2-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="484a2-165">Requirements</span></span>

|<span data-ttu-id="484a2-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="484a2-166">Requirement</span></span>| <span data-ttu-id="484a2-167">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="484a2-168">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="484a2-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484a2-169">1.0</span><span class="sxs-lookup"><span data-stu-id="484a2-169">1.0</span></span>|
|[<span data-ttu-id="484a2-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="484a2-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484a2-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484a2-171">ReadItem</span></span>|
|[<span data-ttu-id="484a2-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="484a2-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484a2-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="484a2-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484a2-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="484a2-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="484a2-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="484a2-175">timeZone :String</span></span>

<span data-ttu-id="484a2-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="484a2-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="484a2-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="484a2-177">Type:</span></span>

*   <span data-ttu-id="484a2-178">String</span><span class="sxs-lookup"><span data-stu-id="484a2-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="484a2-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="484a2-179">Requirements</span></span>

|<span data-ttu-id="484a2-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="484a2-180">Requirement</span></span>| <span data-ttu-id="484a2-181">Valor</span><span class="sxs-lookup"><span data-stu-id="484a2-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="484a2-182">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="484a2-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="484a2-183">1.0</span><span class="sxs-lookup"><span data-stu-id="484a2-183">1.0</span></span>|
|[<span data-ttu-id="484a2-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="484a2-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="484a2-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="484a2-185">ReadItem</span></span>|
|[<span data-ttu-id="484a2-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="484a2-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="484a2-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="484a2-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="484a2-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="484a2-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```