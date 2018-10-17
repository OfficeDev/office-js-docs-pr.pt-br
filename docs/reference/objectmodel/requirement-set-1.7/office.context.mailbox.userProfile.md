
# <a name="userprofile"></a><span data-ttu-id="92f58-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="92f58-101">userProfile</span></span>

### <span data-ttu-id="92f58-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="92f58-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="92f58-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92f58-104">Requirements</span></span>

|<span data-ttu-id="92f58-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="92f58-105">Requirement</span></span>| <span data-ttu-id="92f58-106">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="92f58-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92f58-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92f58-108">1.0</span><span class="sxs-lookup"><span data-stu-id="92f58-108">1.0</span></span>|
|[<span data-ttu-id="92f58-109">Nível mínimo de permissão</span><span class="sxs-lookup"><span data-stu-id="92f58-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92f58-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92f58-110">ReadItem</span></span>|
|[<span data-ttu-id="92f58-111">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="92f58-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92f58-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="92f58-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="92f58-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="92f58-113">Members and methods</span></span>

| <span data-ttu-id="92f58-114">Membro</span><span class="sxs-lookup"><span data-stu-id="92f58-114">Member</span></span> | <span data-ttu-id="92f58-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="92f58-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="92f58-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="92f58-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="92f58-117">Membro</span><span class="sxs-lookup"><span data-stu-id="92f58-117">Member</span></span> |
| [<span data-ttu-id="92f58-118">displayName</span><span class="sxs-lookup"><span data-stu-id="92f58-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="92f58-119">Membro</span><span class="sxs-lookup"><span data-stu-id="92f58-119">Member</span></span> |
| [<span data-ttu-id="92f58-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="92f58-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="92f58-121">Membro</span><span class="sxs-lookup"><span data-stu-id="92f58-121">Member</span></span> |
| [<span data-ttu-id="92f58-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="92f58-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="92f58-123">Membro</span><span class="sxs-lookup"><span data-stu-id="92f58-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="92f58-124">Membros</span><span class="sxs-lookup"><span data-stu-id="92f58-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="92f58-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="92f58-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="92f58-126">Atualmente, este membro só é suportado no Outlook 2016 para Mac, build 16.9.1212 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="92f58-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="92f58-p102">Obtém o tipo de conta do usuário associado com a caixa de correio. Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="92f58-p102">Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.</span></span>

| <span data-ttu-id="92f58-129">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-129">Value</span></span> | <span data-ttu-id="92f58-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="92f58-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="92f58-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="92f58-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="92f58-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="92f58-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="92f58-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="92f58-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="92f58-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="92f58-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="92f58-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="92f58-135">Type:</span></span>

*   <span data-ttu-id="92f58-136">String</span><span class="sxs-lookup"><span data-stu-id="92f58-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92f58-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92f58-137">Requirements</span></span>

|<span data-ttu-id="92f58-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="92f58-138">Requirement</span></span>| <span data-ttu-id="92f58-139">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="92f58-140">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92f58-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92f58-141">1.6</span><span class="sxs-lookup"><span data-stu-id="92f58-141">-16</span></span> |
|[<span data-ttu-id="92f58-142">Nível mínimo de permissão</span><span class="sxs-lookup"><span data-stu-id="92f58-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92f58-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92f58-143">ReadItem</span></span>|
|[<span data-ttu-id="92f58-144">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="92f58-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92f58-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="92f58-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92f58-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92f58-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="92f58-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="92f58-147">displayName :String</span></span>

<span data-ttu-id="92f58-148">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="92f58-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="92f58-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="92f58-149">Type:</span></span>

*   <span data-ttu-id="92f58-150">String</span><span class="sxs-lookup"><span data-stu-id="92f58-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92f58-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92f58-151">Requirements</span></span>

|<span data-ttu-id="92f58-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="92f58-152">Requirement</span></span>| <span data-ttu-id="92f58-153">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="92f58-154">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92f58-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92f58-155">1.0</span><span class="sxs-lookup"><span data-stu-id="92f58-155">1.0</span></span>|
|[<span data-ttu-id="92f58-156">Nível mínimo de permissão</span><span class="sxs-lookup"><span data-stu-id="92f58-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92f58-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92f58-157">ReadItem</span></span>|
|[<span data-ttu-id="92f58-158">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="92f58-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92f58-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="92f58-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92f58-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92f58-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="92f58-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="92f58-161">emailAddress :String</span></span>

<span data-ttu-id="92f58-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="92f58-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="92f58-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="92f58-163">Type:</span></span>

*   <span data-ttu-id="92f58-164">String</span><span class="sxs-lookup"><span data-stu-id="92f58-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92f58-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92f58-165">Requirements</span></span>

|<span data-ttu-id="92f58-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="92f58-166">Requirement</span></span>| <span data-ttu-id="92f58-167">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="92f58-168">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92f58-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92f58-169">1.0</span><span class="sxs-lookup"><span data-stu-id="92f58-169">1.0</span></span>|
|[<span data-ttu-id="92f58-170">Nível mínimo de permissão</span><span class="sxs-lookup"><span data-stu-id="92f58-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92f58-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92f58-171">ReadItem</span></span>|
|[<span data-ttu-id="92f58-172">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="92f58-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92f58-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="92f58-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92f58-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92f58-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="92f58-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="92f58-175">timeZone :String</span></span>

<span data-ttu-id="92f58-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="92f58-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="92f58-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="92f58-177">Type:</span></span>

*   <span data-ttu-id="92f58-178">String</span><span class="sxs-lookup"><span data-stu-id="92f58-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92f58-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92f58-179">Requirements</span></span>

|<span data-ttu-id="92f58-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="92f58-180">Requirement</span></span>| <span data-ttu-id="92f58-181">Valor</span><span class="sxs-lookup"><span data-stu-id="92f58-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="92f58-182">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92f58-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92f58-183">1.0</span><span class="sxs-lookup"><span data-stu-id="92f58-183">1.0</span></span>|
|[<span data-ttu-id="92f58-184">Nível mínimo de permissão</span><span class="sxs-lookup"><span data-stu-id="92f58-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92f58-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92f58-185">ReadItem</span></span>|
|[<span data-ttu-id="92f58-186">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="92f58-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="92f58-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="92f58-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="92f58-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92f58-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```