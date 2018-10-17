
# <a name="userprofile"></a><span data-ttu-id="b4d54-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b4d54-101">userProfile</span></span>

### <span data-ttu-id="b4d54-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b4d54-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4d54-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b4d54-104">Requirements</span></span>

|<span data-ttu-id="b4d54-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b4d54-105">Requirement</span></span>| <span data-ttu-id="b4d54-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4d54-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b4d54-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4d54-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b4d54-108">1.0</span></span>|
|[<span data-ttu-id="b4d54-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b4d54-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4d54-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4d54-110">ReadItem</span></span>|
|[<span data-ttu-id="b4d54-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b4d54-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4d54-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b4d54-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b4d54-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b4d54-113">Members and methods</span></span>

| <span data-ttu-id="b4d54-114">Membro</span><span class="sxs-lookup"><span data-stu-id="b4d54-114">Member</span></span> | <span data-ttu-id="b4d54-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="b4d54-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="b4d54-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="b4d54-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="b4d54-117">Membro</span><span class="sxs-lookup"><span data-stu-id="b4d54-117">Member</span></span> |
| [<span data-ttu-id="b4d54-118">displayName</span><span class="sxs-lookup"><span data-stu-id="b4d54-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="b4d54-119">Membro</span><span class="sxs-lookup"><span data-stu-id="b4d54-119">Member</span></span> |
| [<span data-ttu-id="b4d54-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b4d54-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b4d54-121">Membro</span><span class="sxs-lookup"><span data-stu-id="b4d54-121">Member</span></span> |
| [<span data-ttu-id="b4d54-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="b4d54-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b4d54-123">Membro</span><span class="sxs-lookup"><span data-stu-id="b4d54-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b4d54-124">Membros</span><span class="sxs-lookup"><span data-stu-id="b4d54-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="b4d54-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="b4d54-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="b4d54-126">Esse membro é compatível somente no Outlook 2016 ou posterior para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="b4d54-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="b4d54-p102">Obtém o tipo de conta do usuário associado com a caixa de correio. Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="b4d54-p102">Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.</span></span>

| <span data-ttu-id="b4d54-129">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-129">Value</span></span> | <span data-ttu-id="b4d54-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="b4d54-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="b4d54-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="b4d54-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="b4d54-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="b4d54-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="b4d54-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="b4d54-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="b4d54-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="b4d54-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="b4d54-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b4d54-135">Type:</span></span>

*   <span data-ttu-id="b4d54-136">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="b4d54-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4d54-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b4d54-137">Requirements</span></span>

|<span data-ttu-id="b4d54-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="b4d54-138">Requirement</span></span>| <span data-ttu-id="b4d54-139">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4d54-140">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b4d54-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4d54-141">1.6</span><span class="sxs-lookup"><span data-stu-id="b4d54-141">-16</span></span> |
|[<span data-ttu-id="b4d54-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b4d54-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4d54-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4d54-143">ReadItem</span></span>|
|[<span data-ttu-id="b4d54-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b4d54-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4d54-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b4d54-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4d54-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b4d54-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="b4d54-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b4d54-147">displayName :String</span></span>

<span data-ttu-id="b4d54-148">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="b4d54-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b4d54-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b4d54-149">Type:</span></span>

*   <span data-ttu-id="b4d54-150">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="b4d54-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4d54-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b4d54-151">Requirements</span></span>

|<span data-ttu-id="b4d54-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="b4d54-152">Requirement</span></span>| <span data-ttu-id="b4d54-153">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4d54-154">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b4d54-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4d54-155">1.0</span><span class="sxs-lookup"><span data-stu-id="b4d54-155">1.0</span></span>|
|[<span data-ttu-id="b4d54-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b4d54-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4d54-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4d54-157">ReadItem</span></span>|
|[<span data-ttu-id="b4d54-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b4d54-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4d54-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b4d54-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4d54-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b4d54-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b4d54-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b4d54-161">emailAddress :String</span></span>

<span data-ttu-id="b4d54-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="b4d54-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b4d54-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b4d54-163">Type:</span></span>

*   <span data-ttu-id="b4d54-164">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="b4d54-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4d54-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b4d54-165">Requirements</span></span>

|<span data-ttu-id="b4d54-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="b4d54-166">Requirement</span></span>| <span data-ttu-id="b4d54-167">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4d54-168">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b4d54-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4d54-169">1.0</span><span class="sxs-lookup"><span data-stu-id="b4d54-169">1.0</span></span>|
|[<span data-ttu-id="b4d54-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b4d54-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4d54-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4d54-171">ReadItem</span></span>|
|[<span data-ttu-id="b4d54-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b4d54-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4d54-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b4d54-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4d54-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b4d54-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b4d54-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b4d54-175">timeZone :String</span></span>

<span data-ttu-id="b4d54-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="b4d54-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b4d54-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b4d54-177">Type:</span></span>

*   <span data-ttu-id="b4d54-178">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="b4d54-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4d54-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b4d54-179">Requirements</span></span>

|<span data-ttu-id="b4d54-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="b4d54-180">Requirement</span></span>| <span data-ttu-id="b4d54-181">Valor</span><span class="sxs-lookup"><span data-stu-id="b4d54-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4d54-182">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b4d54-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4d54-183">1.0</span><span class="sxs-lookup"><span data-stu-id="b4d54-183">1.0</span></span>|
|[<span data-ttu-id="b4d54-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b4d54-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4d54-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4d54-185">ReadItem</span></span>|
|[<span data-ttu-id="b4d54-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b4d54-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4d54-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b4d54-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4d54-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b4d54-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```