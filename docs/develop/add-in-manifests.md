---
title: Manifesto XML dos Suplementos do Office
description: Obtenha uma visão geral do manifesto de suplemento do Office e seus usos.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: e664893445ed6d9ee9a7adf23f3b3b189df8634e
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839996"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="e8ad4-103">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-103">Office Add-ins XML manifest</span></span>

<span data-ttu-id="e8ad4-104">O arquivo de manifesto XML de um Suplemento do Office descreve como seu suplemento deve ser ativado quando um usuário final o instala e usa com os aplicativos e documentos do Office.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-104">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="e8ad4-105">Um arquivo de manifesto XML com base nesse esquema permite que um Suplemento do Office faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="e8ad4-105">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="e8ad4-106">Descreva a si mesmo fornecendo ID, versão, descrição, nome para exibição e local padrão.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-106">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="e8ad4-107">Especifique as imagens usadas para identidade visual do suplemento e a iconografia usada para os [comandos do suplemento][] na faixa de opções do Aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-107">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office app ribbon.</span></span>

* <span data-ttu-id="e8ad4-108">Especifique como o suplemento se integra ao Office, incluindo qualquer interface do usuário personalizada, como botões da faixa de opções criados pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-108">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="e8ad4-109">Especifique as dimensões padrão solicitadas para suplementos de conteúdo e a altura solicitada para Suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-109">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="e8ad4-110">Declare permissões exigidas pelo Suplemento do Office, como ler ou gravar no documento.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-110">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="e8ad4-111">Para os suplementos do Outlook, defina a regra ou as regras que especificam o contexto no qual serão ativados e interagirão com uma mensagem, compromisso ou item de solicitação da reunião.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-111">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a><span data-ttu-id="e8ad4-112">Elementos exigidos</span><span class="sxs-lookup"><span data-stu-id="e8ad4-112">Required elements</span></span>

<span data-ttu-id="e8ad4-113">A tabela a seguir especifica os elementos exigidos para os três tipos de Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-113">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="e8ad4-114">Também há uma ordem obrigatória na qual os elementos devem aparecer dentro de seu elemento-pai.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-114">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="e8ad4-115">Confira mais informações em [Como encontrar a ordem adequada de elementos de manifesto](manifest-element-ordering.md).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-115">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="e8ad4-116">Elementos obrigatórios de acordo com o tipo de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-116">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="e8ad4-117">Elemento</span><span class="sxs-lookup"><span data-stu-id="e8ad4-117">Element</span></span>                                                                                      | <span data-ttu-id="e8ad4-118">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e8ad4-118">Content</span></span> | <span data-ttu-id="e8ad4-119">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e8ad4-119">Task pane</span></span> | <span data-ttu-id="e8ad4-120">Outlook</span><span class="sxs-lookup"><span data-stu-id="e8ad4-120">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="e8ad4-121">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-121">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="e8ad4-122">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-122">X</span></span>    |     <span data-ttu-id="e8ad4-123">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-123">X</span></span>     |    <span data-ttu-id="e8ad4-124">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-124">X</span></span>    |
| <span data-ttu-id="e8ad4-125">[Id][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-125">[Id][]</span></span>                                                                                       |    <span data-ttu-id="e8ad4-126">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-126">X</span></span>    |     <span data-ttu-id="e8ad4-127">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-127">X</span></span>     |    <span data-ttu-id="e8ad4-128">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-128">X</span></span>    |
| <span data-ttu-id="e8ad4-129">[Versão][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-129">[Version][]</span></span>                                                                                  |    <span data-ttu-id="e8ad4-130">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-130">X</span></span>    |     <span data-ttu-id="e8ad4-131">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-131">X</span></span>     |    <span data-ttu-id="e8ad4-132">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-132">X</span></span>    |
| <span data-ttu-id="e8ad4-133">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-133">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="e8ad4-134">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-134">X</span></span>    |     <span data-ttu-id="e8ad4-135">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-135">X</span></span>     |    <span data-ttu-id="e8ad4-136">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-136">X</span></span>    |
| <span data-ttu-id="e8ad4-137">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-137">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="e8ad4-138">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-138">X</span></span>    |     <span data-ttu-id="e8ad4-139">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-139">X</span></span>     |    <span data-ttu-id="e8ad4-140">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-140">X</span></span>    |
| <span data-ttu-id="e8ad4-141">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-141">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="e8ad4-142">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-142">X</span></span>    |     <span data-ttu-id="e8ad4-143">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-143">X</span></span>     |    <span data-ttu-id="e8ad4-144">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-144">X</span></span>    |
| <span data-ttu-id="e8ad4-145">[Descrição][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-145">[Description][]</span></span>                                                                              |    <span data-ttu-id="e8ad4-146">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-146">X</span></span>    |     <span data-ttu-id="e8ad4-147">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-147">X</span></span>     |    <span data-ttu-id="e8ad4-148">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-148">X</span></span>    |
| <span data-ttu-id="e8ad4-149">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-149">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="e8ad4-150">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-150">X</span></span>    |     <span data-ttu-id="e8ad4-151">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-151">X</span></span>     |    <span data-ttu-id="e8ad4-152">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-152">X</span></span>    |
| <span data-ttu-id="e8ad4-153">[SupportUrl][]\*\*</span><span class="sxs-lookup"><span data-stu-id="e8ad4-153">[SupportUrl][]\*\*</span></span>                                                                           |    <span data-ttu-id="e8ad4-154">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-154">X</span></span>    |     <span data-ttu-id="e8ad4-155">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-155">X</span></span>     |    <span data-ttu-id="e8ad4-156">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-156">X</span></span>    |
| <span data-ttu-id="e8ad4-157">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-157">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="e8ad4-158">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-158">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="e8ad4-159">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-159">X</span></span>    |     <span data-ttu-id="e8ad4-160">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-160">X</span></span>     |         |
| <span data-ttu-id="e8ad4-161">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-161">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="e8ad4-162">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-162">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="e8ad4-163">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-163">X</span></span>    |     <span data-ttu-id="e8ad4-164">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-164">X</span></span>     |         |
| <span data-ttu-id="e8ad4-165">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-165">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="e8ad4-166">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-166">X</span></span>    |
| <span data-ttu-id="e8ad4-167">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-167">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="e8ad4-168">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-168">X</span></span>    |
| <span data-ttu-id="e8ad4-169">[Permissões (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-169">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="e8ad4-170">[Permissões (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-170">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="e8ad4-171">[Permissões (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-171">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="e8ad4-172">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-172">X</span></span>    |     <span data-ttu-id="e8ad4-173">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-173">X</span></span>     |    <span data-ttu-id="e8ad4-174">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-174">X</span></span>    |
| <span data-ttu-id="e8ad4-175">[Regra (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-175">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="e8ad4-176">[Regra (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-176">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="e8ad4-177">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-177">X</span></span>    |
| <span data-ttu-id="e8ad4-178">[Requisitos (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-178">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="e8ad4-179">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-179">X</span></span>    |
| <span data-ttu-id="e8ad4-180">[Conjunto\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-180">[Set\*][]</span></span><br/><span data-ttu-id="e8ad4-181">[Conjuntos (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-181">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="e8ad4-182">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-182">X</span></span>    |
| <span data-ttu-id="e8ad4-183">[Formulário\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-183">[Form\*][]</span></span><br/><span data-ttu-id="e8ad4-184">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-184">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="e8ad4-185">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-185">X</span></span>    |
| <span data-ttu-id="e8ad4-186">[Conjuntos (Requisitos)\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-186">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="e8ad4-187">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-187">X</span></span>    |     <span data-ttu-id="e8ad4-188">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-188">X</span></span>     |         |
| <span data-ttu-id="e8ad4-189">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-189">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="e8ad4-190">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-190">X</span></span>    |     <span data-ttu-id="e8ad4-191">X</span><span class="sxs-lookup"><span data-stu-id="e8ad4-191">X</span></span>     |         |

<span data-ttu-id="e8ad4-192">_\*Adicionados no esquema de manifesto de suplementos da versão 1.1 do Office._</span><span class="sxs-lookup"><span data-stu-id="e8ad4-192">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<span data-ttu-id="e8ad4-193">_\*\* SupportUrl só é necessário para suplementos distribuídos pelo AppSource._</span><span class="sxs-lookup"><span data-stu-id="e8ad4-193">_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._</span></span>

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissões (contentapp)]: ../reference/manifest/permissions.md
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissões (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissões (mailapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[regra (rulecollection)]: ../reference/manifest/rule.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[regra (mailapp)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[requisitos (mailapp)]: ../reference/manifest/requirements.md
[requirements (mailapp)\*]: ../reference/manifest/requirements.md
[set*]: ../reference/manifest/set.md
[conjuntos (mailapprequirements)\*]: ../reference/manifest/sets.md
[sets (mailapprequirements)\*]: ../reference/manifest/sets.md
[formulário\*]: ../reference/manifest/form.md
[form\*]: ../reference/manifest/form.md
[formsettings*]: ../reference/manifest/formsettings.md
[conjuntos (requisitos)\*]: ../reference/manifest/sets.md
[sets (requirements)\*]: ../reference/manifest/sets.md
[hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a><span data-ttu-id="e8ad4-221">Requisitos de hospedagem</span><span class="sxs-lookup"><span data-stu-id="e8ad4-221">Hosting requirements</span></span>

<span data-ttu-id="e8ad4-222">Todas as imagem URIs, como as usadas para os [comandos do suplemento][], devem ser compatíveis com armazenamento em cache.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-222">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="e8ad4-223">O servidor que hospeda a imagem não deve retornar um cabeçalho `Cache-Control` especificando `no-cache`, `no-store` ou opções semelhantes na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-223">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="e8ad4-224">Todas as URLs, como os locais dos arquivos de origem especificados no elemento [SourceLocation](../reference/manifest/sourcelocation.md), devem estar **protegidos por SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-224">All URLs, such as the source file locations specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="e8ad4-225">Práticas recomendadas de envio ao AppSource</span><span class="sxs-lookup"><span data-stu-id="e8ad4-225">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="e8ad4-p103">Verifique se a identificação do suplemento é um GUID válido e exclusivo. Diversas ferramentas de gerador de GUID estão disponíveis na Web e podem ser usadas para criar um GUID exclusivo.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="e8ad4-228">Os suplementos enviados ao AppSource também devem conter o elemento [SupportUrl](../reference/manifest/supporturl.md).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-228">Add-ins submitted to AppSource must also include the [SupportUrl](../reference/manifest/supporturl.md) element.</span></span> <span data-ttu-id="e8ad4-229">Saiba mais em [Políticas de validação para aplicativos e suplementos enviados ao AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-229">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/legal/marketplace/certification-policies).</span></span>

<span data-ttu-id="e8ad4-230">Use apenas o elemento [AppDomains](../reference/manifest/appdomains.md) para especificar domínios diferentes daqueles especificados no elemento [SourceLocation](../reference/manifest/sourcelocation.md) para cenários de autenticação.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-230">Only use the [AppDomains](../reference/manifest/appdomains.md) element to specify domains other than the one specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="e8ad4-231">Especificar os domínios que você deseja abrir na janela do suplemento</span><span class="sxs-lookup"><span data-stu-id="e8ad4-231">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="e8ad4-232">Ao executar no Office Online, o seu painel de tarefas pode ser navegado para qualquer URL.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-232">When running in Office on the web, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="e8ad4-233">No entanto, nas plataformas de desktop, se o suplemento tentar acessar uma URL em um domínio diferente do domínio que hospeda a página inicial (conforme especificado no elemento [SourceLocation](../reference/manifest/sourcelocation.md) do arquivo de manifesto), essa URL abre em uma nova janela de navegador fora do painel de suplementos do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-233">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.</span></span>

<span data-ttu-id="e8ad4-234">Para substituir esse comportamento (Office para desktop), especifique cada domínio que você deseja abrir na janela do suplemento na lista de domínios especificados no elemento [AppDomains](../reference/manifest/appdomains.md) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-234">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="e8ad4-235">Se o suplemento tentar ir para uma URL em um domínio que está na lista, ela então abre no painel de tarefas do Office para desktop e no Office Online.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-235">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop.</span></span> <span data-ttu-id="e8ad4-236">Se ele tentar acessar uma URL que não está na lista, no Office para desktop, essa URL abre em uma nova janela do navegador (fora do painel de suplementos).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-236">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="e8ad4-237">Há duas exceções para esse comportamento:</span><span class="sxs-lookup"><span data-stu-id="e8ad4-237">There are two exceptions to this behavior:</span></span>
>
> - <span data-ttu-id="e8ad4-238">Isso se aplica somente ao painel raiz do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-238">It applies only to the root pane of the add-in.</span></span> <span data-ttu-id="e8ad4-239">Se houver um iframe inserido na página do suplemento, o iframe pode ser direcionado para qualquer URL independentemente se ele está listado na **AppDomains**, até mesmo no Office para desktop.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-239">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="e8ad4-240">Quando uma caixa de diálogo é aberta coma API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#displaydialogasync-startaddress--options--callback-), a URL que é passada para o método deve estar no mesmo domínio do suplemento, mas a caixa de diálogo pode ser direcionada para qualquer URL, independentemente de estar listada no **AppDomains**, mesmo no Office para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-240">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js&preserve-view=true#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="e8ad4-241">O exemplo de manifesto XML a seguir hospeda sua página de suplemento principal no domínio `https://www.contoso.com`, conforme especificado no elemento **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-241">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="e8ad4-242">Ele também especifica o domínio `https://www.northwindtraders.com` em um elemento [AppDomain](../reference/manifest/appdomain.md), dentro da lista de elementos **AppDomains**</span><span class="sxs-lookup"><span data-stu-id="e8ad4-242">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](../reference/manifest/appdomain.md) element within the **AppDomains** element list.</span></span> <span data-ttu-id="e8ad4-243">Se o suplemento acessar uma página no `www.northwindtraders.com`domínio, essa página abre no painel do suplemento, mesmo na área de trabalho do Office.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-243">If the add-in goes to a page in the `www.northwindtraders.com` domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="e8ad4-244">Especificar domínios a partir dos quais as chamadas da API do Office.js são feitas</span><span class="sxs-lookup"><span data-stu-id="e8ad4-244">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="e8ad4-245">Seu suplemento pode fazer chamadas API do Office.js a partir do domínio referenciado no elemento[SourceLocation](../reference/manifest/sourcelocation.md) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-245">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file.</span></span> <span data-ttu-id="e8ad4-246">Se você tiver outros iFrames dentro de seu suplemento que precisem acessar APIs do Office.js, adicione o domínio dessa URL de origem à lista especificada no elemento [AppDomains](../reference/manifest/appdomains.md) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-246">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="e8ad4-247">Se um iFrame com uma fonte não incluída na lista `AppDomains` tentar fazer uma chamada de API do Office. js, o suplemento receberá um[ erro de permissão negada](../reference/javascript-api-for-office-error-codes.md).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-247">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span>

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="e8ad4-248">Exemplos e esquemas do arquivo XML de manifesto v1.1</span><span class="sxs-lookup"><span data-stu-id="e8ad4-248">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="e8ad4-249">As seções a seguir mostram exemplos de arquivos XML de manifesto v1.1 para suplementos de conteúdo, de painel de tarefas e do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e8ad4-249">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-pane"></a>[<span data-ttu-id="e8ad4-250">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e8ad4-250">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="e8ad4-251">Esquemas de manifesto de suplemento</span><span class="sxs-lookup"><span data-stu-id="e8ad4-251">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office app ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://myCDN/Images/ButtonFunction.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="content"></a>[<span data-ttu-id="e8ad4-252">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e8ad4-252">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="e8ad4-253">Esquemas de manifesto de suplemento</span><span class="sxs-lookup"><span data-stu-id="e8ad4-253">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mail"></a>[<span data-ttu-id="e8ad4-254">Email</span><span class="sxs-lookup"><span data-stu-id="e8ad4-254">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="e8ad4-255">Esquemas de manifesto de suplemento</span><span class="sxs-lookup"><span data-stu-id="e8ad4-255">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="e8ad4-256">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-256">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="e8ad4-257">Para saber mais sobre como validar um manifesto em relação à [Definição do Esquema XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), confira [Validar o manifesto de suplemento do Office](../testing/troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="e8ad4-257">For information about validating a manifest against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e8ad4-258">Confira também</span><span class="sxs-lookup"><span data-stu-id="e8ad4-258">See also</span></span>

* [<span data-ttu-id="e8ad4-259">Como identificar a ordem correta dos elementos do manifesto</span><span class="sxs-lookup"><span data-stu-id="e8ad4-259">How to find the proper order of manifest elements</span></span>](manifest-element-ordering.md)
* <span data-ttu-id="e8ad4-260">[Criar comandos de suplementos em seu manifesto][comandos de suplementos]</span><span class="sxs-lookup"><span data-stu-id="e8ad4-260">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="e8ad4-261">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-261">Specify Office applications and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="e8ad4-262">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-262">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="e8ad4-263">Referência de esquema para manifestos de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-263">Schema reference for Office Add-ins manifests</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [<span data-ttu-id="e8ad4-264">Atualizar a versão da API e do manifesto</span><span class="sxs-lookup"><span data-stu-id="e8ad4-264">Update API and manifest version</span></span>](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [<span data-ttu-id="e8ad4-265">Identificar um suplemento COM equivalente</span><span class="sxs-lookup"><span data-stu-id="e8ad4-265">Identify an equivalent COM add-in</span></span>](make-office-add-in-compatible-with-existing-com-add-in.md)
* [<span data-ttu-id="e8ad4-266">Solicitar permissões para uso da API em suplementos </span><span class="sxs-lookup"><span data-stu-id="e8ad4-266">Requesting permissions for API use in add-ins</span></span>](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [<span data-ttu-id="e8ad4-267">Validar o manifesto de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="e8ad4-267">Validate an Office Add-in's manifest</span></span>](../testing/troubleshoot-manifest.md)

[Comandos de suplemento]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
