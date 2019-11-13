---
title: Manifesto XML dos Suplementos do Office
description: ''
ms.date: 09/26/2019
localization_priority: Priority
ms.openlocfilehash: 9a3081c13756bee668081e1fb9d04cf5458320d4
ms.sourcegitcommit: 88d81aa2d707105cf0eb55d9774b2e7cf468b03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/13/2019
ms.locfileid: "38301971"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="c3e5f-102">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c3e5f-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="c3e5f-103">O arquivo de manifesto XML de um Suplemento do Office descreve como seu suplemento deve ser ativado quando um usuário final o instala e usa com os aplicativos e documentos do Office.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="c3e5f-104">Um arquivo de manifesto XML com base nesse esquema permite que um Suplemento do Office faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="c3e5f-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="c3e5f-105">Descreva a si mesmo fornecendo ID, versão, descrição, nome para exibição e local padrão.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="c3e5f-106">Especifique as imagens usadas para identidade visual do suplemento e a iconografia usada para os [comandos do suplemento][] na Faixa de Opções do Office.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-106">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="c3e5f-107">Especifique como o suplemento se integra ao Office, incluindo qualquer interface do usuário personalizada, como botões da faixa de opções criados pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="c3e5f-108">Especifique as dimensões padrão solicitadas para suplementos de conteúdo e a altura solicitada para Suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="c3e5f-109">Declare permissões exigidas pelo Suplemento do Office, como ler ou gravar no documento.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="c3e5f-110">Para os suplementos do Outlook, defina a regra ou as regras que especificam o contexto no qual serão ativados e interagirão com uma mensagem, compromisso ou item de solicitação da reunião.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="c3e5f-p101">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="c3e5f-113">Elementos exigidos</span><span class="sxs-lookup"><span data-stu-id="c3e5f-113">Required elements</span></span>

<span data-ttu-id="c3e5f-114">A tabela a seguir especifica os elementos exigidos para os três tipos de Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="c3e5f-115">Também há uma ordem obrigatória na qual os elementos devem aparecer dentro de seu elemento-pai.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-115">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="c3e5f-116">Confira mais informações em [Como encontrar a ordem adequada de elementos de manifesto](manifest-element-ordering.md).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-116">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="c3e5f-117">Elementos obrigatórios de acordo com o tipo de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="c3e5f-117">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="c3e5f-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="c3e5f-118">Element</span></span>                                                                                      | <span data-ttu-id="c3e5f-119">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3e5f-119">Content</span></span> | <span data-ttu-id="c3e5f-120">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="c3e5f-120">Task pane</span></span> | <span data-ttu-id="c3e5f-121">Outlook</span><span class="sxs-lookup"><span data-stu-id="c3e5f-121">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="c3e5f-122">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-122">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="c3e5f-123">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-123">X</span></span>    |     <span data-ttu-id="c3e5f-124">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-124">X</span></span>     |    <span data-ttu-id="c3e5f-125">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-125">X</span></span>    |
| <span data-ttu-id="c3e5f-126">[Id][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-126">[Id][]</span></span>                                                                                       |    <span data-ttu-id="c3e5f-127">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-127">X</span></span>    |     <span data-ttu-id="c3e5f-128">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-128">X</span></span>     |    <span data-ttu-id="c3e5f-129">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-129">X</span></span>    |
| <span data-ttu-id="c3e5f-130">[Versão][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-130">[Version][]</span></span>                                                                                  |    <span data-ttu-id="c3e5f-131">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-131">X</span></span>    |     <span data-ttu-id="c3e5f-132">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-132">X</span></span>     |    <span data-ttu-id="c3e5f-133">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-133">X</span></span>    |
| <span data-ttu-id="c3e5f-134">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-134">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="c3e5f-135">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-135">X</span></span>    |     <span data-ttu-id="c3e5f-136">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-136">X</span></span>     |    <span data-ttu-id="c3e5f-137">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-137">X</span></span>    |
| <span data-ttu-id="c3e5f-138">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-138">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="c3e5f-139">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-139">X</span></span>    |     <span data-ttu-id="c3e5f-140">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-140">X</span></span>     |    <span data-ttu-id="c3e5f-141">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-141">X</span></span>    |
| <span data-ttu-id="c3e5f-142">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-142">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="c3e5f-143">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-143">X</span></span>    |     <span data-ttu-id="c3e5f-144">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-144">X</span></span>     |    <span data-ttu-id="c3e5f-145">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-145">X</span></span>    |
| <span data-ttu-id="c3e5f-146">[Descrição][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-146">[Description][]</span></span>                                                                              |    <span data-ttu-id="c3e5f-147">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-147">X</span></span>    |     <span data-ttu-id="c3e5f-148">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-148">X</span></span>     |    <span data-ttu-id="c3e5f-149">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-149">X</span></span>    |
| <span data-ttu-id="c3e5f-150">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-150">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="c3e5f-151">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-151">X</span></span>    |     <span data-ttu-id="c3e5f-152">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-152">X</span></span>     |    <span data-ttu-id="c3e5f-153">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-153">X</span></span>    |
| <span data-ttu-id="c3e5f-154">[SupportUrl][]\*\*</span><span class="sxs-lookup"><span data-stu-id="c3e5f-154">[SupportUrl][]\*\*</span></span>                                                                           |    <span data-ttu-id="c3e5f-155">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-155">X</span></span>    |     <span data-ttu-id="c3e5f-156">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-156">X</span></span>     |    <span data-ttu-id="c3e5f-157">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-157">X</span></span>    |
| <span data-ttu-id="c3e5f-158">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-158">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="c3e5f-159">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-159">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="c3e5f-160">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-160">X</span></span>    |     <span data-ttu-id="c3e5f-161">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-161">X</span></span>     |         |
| <span data-ttu-id="c3e5f-162">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-162">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="c3e5f-163">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-163">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="c3e5f-164">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-164">X</span></span>    |     <span data-ttu-id="c3e5f-165">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-165">X</span></span>     |         |
| <span data-ttu-id="c3e5f-166">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-166">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="c3e5f-167">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-167">X</span></span>    |
| <span data-ttu-id="c3e5f-168">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-168">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="c3e5f-169">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-169">X</span></span>    |
| <span data-ttu-id="c3e5f-170">[Permissões (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-170">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="c3e5f-171">[Permissões (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-171">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="c3e5f-172">[Permissões (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-172">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="c3e5f-173">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-173">X</span></span>    |     <span data-ttu-id="c3e5f-174">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-174">X</span></span>     |    <span data-ttu-id="c3e5f-175">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-175">X</span></span>    |
| <span data-ttu-id="c3e5f-176">[Regra (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-176">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="c3e5f-177">[Regra (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-177">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="c3e5f-178">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-178">X</span></span>    |
| <span data-ttu-id="c3e5f-179">[Requisitos (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-179">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="c3e5f-180">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-180">X</span></span>    |
| <span data-ttu-id="c3e5f-181">[Conjunto\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-181">[Set\*][]</span></span><br/><span data-ttu-id="c3e5f-182">[Conjuntos (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-182">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="c3e5f-183">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-183">X</span></span>    |
| <span data-ttu-id="c3e5f-184">[Formulário\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-184">[Form\*][]</span></span><br/><span data-ttu-id="c3e5f-185">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-185">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="c3e5f-186">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-186">X</span></span>    |
| <span data-ttu-id="c3e5f-187">[Conjuntos (Requisitos)\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-187">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="c3e5f-188">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-188">X</span></span>    |     <span data-ttu-id="c3e5f-189">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-189">X</span></span>     |         |
| <span data-ttu-id="c3e5f-190">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-190">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="c3e5f-191">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-191">X</span></span>    |     <span data-ttu-id="c3e5f-192">X</span><span class="sxs-lookup"><span data-stu-id="c3e5f-192">X</span></span>     |         |

<span data-ttu-id="c3e5f-193">_\*Adicionados no esquema de manifesto de suplementos da versão 1.1 do Office._</span><span class="sxs-lookup"><span data-stu-id="c3e5f-193">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<span data-ttu-id="c3e5f-194">_\*\* SupportUrl só é necessário para suplementos distribuídos pelo AppSource._</span><span class="sxs-lookup"><span data-stu-id="c3e5f-194">_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._</span></span>

<!-- Links for above table -->

[officeapp]: /office/dev/add-ins/reference/manifest/officeapp
[id]: /office/dev/add-ins/reference/manifest/id
[version]: /office/dev/add-ins/reference/manifest/version
[providername]: /office/dev/add-ins/reference/manifest/providername
[defaultlocale]: /office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: /office/dev/add-ins/reference/manifest/displayname
[description]: /office/dev/add-ins/reference/manifest/description
[iconurl]: /office/dev/add-ins/reference/manifest/iconurl
[supporturl]: /office/dev/add-ins/reference/manifest/supporturl
[defaultsettings (contentapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: /office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: /office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: https://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissões (contentapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (contentapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissões (taskpaneapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissões (mailapp)]: /office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: /office/dev/add-ins/reference/manifest/permissions
[regra (rulecollection)]: /office/dev/add-ins/reference/manifest/rule
[rule (rulecollection)]: /office/dev/add-ins/reference/manifest/rule
[regra (mailapp)]: /office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: /office/dev/add-ins/reference/manifest/rule
[requisitos (mailapp)]: /office/dev/add-ins/reference/manifest/requirements
[requirements (mailapp)\*]: /office/dev/add-ins/reference/manifest/requirements
[set*]: /office/dev/add-ins/reference/manifest/set
[conjuntos (mailapprequirements)\*]: /office/dev/add-ins/reference/manifest/sets
[sets (mailapprequirements)\*]: /office/dev/add-ins/reference/manifest/sets
[formulário\*]: /office/dev/add-ins/reference/manifest/form
[form\*]: /office/dev/add-ins/reference/manifest/form
[formsettings*]: /office/dev/add-ins/reference/manifest/formsettings
[conjuntos (requisitos)\*]: /office/dev/add-ins/reference/manifest/sets
[sets (requirements)\*]: /office/dev/add-ins/reference/manifest/sets
[hosts*]: /office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="c3e5f-222">Requisitos de hospedagem</span><span class="sxs-lookup"><span data-stu-id="c3e5f-222">Hosting requirements</span></span>

<span data-ttu-id="c3e5f-223">Todas as imagem URIs, como as usadas para os [comandos do suplemento][], devem ser compatíveis com armazenamento em cache.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-223">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="c3e5f-224">O servidor que hospeda a imagem não deve retornar um cabeçalho `Cache-Control` especificando `no-cache`, `no-store` ou opções semelhantes na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-224">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="c3e5f-225">Todas as URLs, como os locais dos arquivos de origem especificados no elemento [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation), devem estar **protegidos por SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-225">All URLs, such as the source file locations specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="c3e5f-226">Práticas recomendadas de envio ao AppSource</span><span class="sxs-lookup"><span data-stu-id="c3e5f-226">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="c3e5f-p104">Verifique se a identificação do suplemento é um GUID válido e exclusivo. Diversas ferramentas de gerador de GUID estão disponíveis na Web e podem ser usadas para criar um GUID exclusivo.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-p104">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="c3e5f-229">Os suplementos enviados ao AppSource também devem conter o elemento [SupportUrl](/office/dev/add-ins/reference/manifest/supporturl).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-229">Add-ins submitted to AppSource must also include the [SupportUrl](/office/dev/add-ins/reference/manifest/supporturl) element.</span></span> <span data-ttu-id="c3e5f-230">Saiba mais em [Políticas de validação para aplicativos e suplementos enviados ao AppSource](/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-230">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="c3e5f-231">Use apenas o elemento [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) para especificar domínios diferentes daqueles especificados no elemento [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) para cenários de autenticação.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-231">Only use the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="c3e5f-232">Especificar os domínios que você deseja abrir na janela do suplemento</span><span class="sxs-lookup"><span data-stu-id="c3e5f-232">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="c3e5f-233">Ao executar no Office Online, o seu painel de tarefas pode ser navegado para qualquer URL.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-233">When running in Office on the web, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="c3e5f-234">No entanto, nas plataformas de desktop, se o suplemento tentar acessar uma URL em um domínio diferente do domínio que hospeda a página inicial (conforme especificado no elemento [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) do arquivo de manifesto), essa URL abre em uma nova janela de navegador fora do painel de suplementos do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-234">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="c3e5f-235">Para substituir esse comportamento (Office para desktop), especifique cada domínio que você deseja abrir na janela do suplemento na lista de domínios especificados no elemento [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-235">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="c3e5f-236">Se o suplemento tentar ir para uma URL em um domínio que está na lista, ela então abre no painel de tarefas do Office para desktop e no Office Online.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-236">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop.</span></span> <span data-ttu-id="c3e5f-237">Se ele tentar acessar uma URL que não está na lista, no Office para desktop, essa URL abre em uma nova janela do navegador (fora do painel de suplementos).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-237">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="c3e5f-238">Há duas exceções para esse comportamento:</span><span class="sxs-lookup"><span data-stu-id="c3e5f-238">There are two exceptions to this behavior:</span></span>
> 
> - <span data-ttu-id="c3e5f-239">Isso se aplica somente ao painel raiz do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-239">It applies only to the root pane of the add-in.</span></span> <span data-ttu-id="c3e5f-240">Se houver um iframe inserido na página do suplemento, o iframe pode ser direcionado para qualquer URL independentemente se ele está listado na **AppDomains**, até mesmo no Office para desktop.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-240">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="c3e5f-241">Quando uma caixa de diálogo é aberta coma API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-), a URL que é passada para o método deve estar no mesmo domínio do suplemento, mas a caixa de diálogo pode ser direcionada para qualquer URL, independentemente de estar listada no **AppDomains**, mesmo no Office para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-241">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span> 

<span data-ttu-id="c3e5f-242">O exemplo de manifesto XML a seguir hospeda sua página de suplemento principal no domínio `https://www.contoso.com`, conforme especificado no elemento **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-242">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="c3e5f-243">Ele também especifica o domínio `https://www.northwindtraders.com` em um elemento [AppDomain](/office/dev/add-ins/reference/manifest/appdomain), dentro da lista de elementos **AppDomains**</span><span class="sxs-lookup"><span data-stu-id="c3e5f-243">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="c3e5f-244">Se o suplemento acessar uma página no domínio www.northwindtraders.com, essa página abre no painel do suplemento, mesmo no Office para desktop.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-244">If the add-in goes to a page in the www.northwindtraders.com domain, that page opens in the add-in pane, even in Office desktop.</span></span>

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

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="c3e5f-245">Especificar domínios a partir dos quais as chamadas da API do Office.js são feitas</span><span class="sxs-lookup"><span data-stu-id="c3e5f-245">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="c3e5f-246">Seu suplemento pode fazer chamadas API do Office.js a partir do domínio referenciado no elemento[SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-246">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file.</span></span> <span data-ttu-id="c3e5f-247">Se você tiver outros iFrames dentro de seu suplemento que precisem acessar APIs do Office.js, adicione o domínio dessa URL de origem à lista especificada no elemento [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-247">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="c3e5f-248">Se um iFrame com uma fonte não incluída na lista `AppDomains` tentar fazer uma chamada de API do Office. js, o suplemento receberá um[ erro de permissão negada](../reference/javascript-api-for-office-error-codes.md).</span><span class="sxs-lookup"><span data-stu-id="c3e5f-248">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span> 

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="c3e5f-249">Exemplos e esquemas do arquivo XML de manifesto v1.1</span><span class="sxs-lookup"><span data-stu-id="c3e5f-249">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="c3e5f-250">As seções a seguir mostram exemplos de arquivos XML de manifesto v1.1 para suplementos de conteúdo, de painel de tarefas e do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-250">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="c3e5f-251">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="c3e5f-251">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="c3e5f-252">Esquema de manifesto do aplicativo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="c3e5f-252">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

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

          <!--PrimaryCommandSurface==Main Office Ribbon-->
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
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
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
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="c3e5f-253">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3e5f-253">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="c3e5f-254">Esquema de manifesto do aplicativo de conteúdo</span><span class="sxs-lookup"><span data-stu-id="c3e5f-254">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="c3e5f-255">Email</span><span class="sxs-lookup"><span data-stu-id="c3e5f-255">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="c3e5f-256">Esquema de manifesto do aplicativo de email</span><span class="sxs-lookup"><span data-stu-id="c3e5f-256">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="c3e5f-257">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="c3e5f-257">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="c3e5f-p111">Para solucionar problemas com seu manifesto, confira [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). Lá, você encontrará informações sobre como validar o manifesto em relação à [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) e também como usar o log de tempo de execução para depurar o manifesto.</span><span class="sxs-lookup"><span data-stu-id="c3e5f-p111">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="c3e5f-260">Confira também</span><span class="sxs-lookup"><span data-stu-id="c3e5f-260">See also</span></span>

* <span data-ttu-id="c3e5f-261">[Criar comandos de suplementos em seu manifesto][comandos de suplementos]</span><span class="sxs-lookup"><span data-stu-id="c3e5f-261">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="c3e5f-262">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="c3e5f-262">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="c3e5f-263">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c3e5f-263">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="c3e5f-264">Referência de esquema para manifestos de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c3e5f-264">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="c3e5f-265">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="c3e5f-265">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[Comandos de suplemento]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
