---
title: Manifesto XML dos Suplementos do Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: b2e0db2712ecfcd9e7df740548968c91ff1c1af2
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004984"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="e3c4f-102">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e3c4f-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="e3c4f-103">O arquivo de manifesto XML de um Suplemento do Office descreve como seu suplemento deve ser ativado quando um usuário final o instala e usa com os aplicativos e documentos do Office.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="e3c4f-104">Um arquivo de manifesto XML com base nesse esquema permite que um Suplemento do Office faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="e3c4f-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="e3c4f-105">Descreva a si mesmo fornecendo ID, versão, descrição, nome para exibição e local padrão.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="e3c4f-106">Especifique as imagens usadas para identidade visual do suplemento e a iconografia usada para os [comandos do suplemento][] na Faixa de Opções do Office.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-106">Specify the images used for branding the Add-in and iconography used for [Add-in Commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="e3c4f-107">Especifique como o suplemento se integra ao Office, incluindo qualquer interface do usuário personalizada, como botões da faixa de opções criados pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="e3c4f-108">Especifique as dimensões padrão solicitadas para suplementos de conteúdo e a altura solicitada para Suplementos do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="e3c4f-109">Declare permissões exigidas pelo Suplemento do Office, como ler ou gravar no documento.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="e3c4f-110">Para os suplementos do Outlook, defina a regra ou as regras que especificam o contexto no qual serão ativados e interagirão com uma mensagem, compromisso ou item de solicitação da reunião.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="e3c4f-p101">Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="e3c4f-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="e3c4f-113">Elementos exigidos</span><span class="sxs-lookup"><span data-stu-id="e3c4f-113">Required elements</span></span>

<span data-ttu-id="e3c4f-114">A tabela a seguir especifica os elementos exigidos para os três tipos de Suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="e3c4f-115">Elementos obrigatórios de acordo com o tipo de Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="e3c4f-115">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="e3c4f-116">Elemento</span><span class="sxs-lookup"><span data-stu-id="e3c4f-116">Element</span></span>                                                                                      | <span data-ttu-id="e3c4f-117">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e3c4f-117">Content</span></span> | <span data-ttu-id="e3c4f-118">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e3c4f-118">Task pane</span></span> | <span data-ttu-id="e3c4f-119">Outlook</span><span class="sxs-lookup"><span data-stu-id="e3c4f-119">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="e3c4f-120">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-120">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="e3c4f-121">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-121">X</span></span>    |     <span data-ttu-id="e3c4f-122">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-122">X</span></span>     |    <span data-ttu-id="e3c4f-123">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-123">X</span></span>    |
| <span data-ttu-id="e3c4f-124">[Id][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-124">[Id][]</span></span>                                                                                       |    <span data-ttu-id="e3c4f-125">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-125">X</span></span>    |     <span data-ttu-id="e3c4f-126">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-126">X</span></span>     |    <span data-ttu-id="e3c4f-127">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-127">X</span></span>    |
| <span data-ttu-id="e3c4f-128">[Versão][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-128">[Version][]</span></span>                                                                                  |    <span data-ttu-id="e3c4f-129">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-129">X</span></span>    |     <span data-ttu-id="e3c4f-130">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-130">X</span></span>     |    <span data-ttu-id="e3c4f-131">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-131">X</span></span>    |
| <span data-ttu-id="e3c4f-132">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-132">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="e3c4f-133">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-133">X</span></span>    |     <span data-ttu-id="e3c4f-134">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-134">X</span></span>     |    <span data-ttu-id="e3c4f-135">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-135">X</span></span>    |
| <span data-ttu-id="e3c4f-136">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-136">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="e3c4f-137">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-137">X</span></span>    |     <span data-ttu-id="e3c4f-138">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-138">X</span></span>     |    <span data-ttu-id="e3c4f-139">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-139">X</span></span>    |
| <span data-ttu-id="e3c4f-140">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-140">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="e3c4f-141">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-141">X</span></span>    |     <span data-ttu-id="e3c4f-142">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-142">X</span></span>     |    <span data-ttu-id="e3c4f-143">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-143">X</span></span>    |
| <span data-ttu-id="e3c4f-144">[Descrição][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-144">[Description][]</span></span>                                                                              |    <span data-ttu-id="e3c4f-145">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-145">X</span></span>    |     <span data-ttu-id="e3c4f-146">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-146">X</span></span>     |    <span data-ttu-id="e3c4f-147">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-147">X</span></span>    |
| <span data-ttu-id="e3c4f-148">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-148">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="e3c4f-149">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-149">X</span></span>    |     <span data-ttu-id="e3c4f-150">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-150">X</span></span>     |    <span data-ttu-id="e3c4f-151">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-151">X</span></span>    |
| <span data-ttu-id="e3c4f-152">[HighResolutionIconUrl][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-152">[HighResolutionIconUrl][]</span></span>                                                                    |    <span data-ttu-id="e3c4f-153">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-153">X</span></span>    |     <span data-ttu-id="e3c4f-154">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-154">X</span></span>     |    <span data-ttu-id="e3c4f-155">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-155">X</span></span>    |
| <span data-ttu-id="e3c4f-156">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-156">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="e3c4f-157">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-157">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="e3c4f-158">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-158">X</span></span>    |     <span data-ttu-id="e3c4f-159">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-159">X</span></span>     |         |
| <span data-ttu-id="e3c4f-160">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-160">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="e3c4f-161">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-161">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="e3c4f-162">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-162">X</span></span>    |     <span data-ttu-id="e3c4f-163">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-163">X</span></span>     |         |
| <span data-ttu-id="e3c4f-164">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-164">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="e3c4f-165">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-165">X</span></span>    |
| <span data-ttu-id="e3c4f-166">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-166">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="e3c4f-167">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-167">X</span></span>    |
| <span data-ttu-id="e3c4f-168">[Permissões (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-168">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="e3c4f-169">[Permissões (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-169">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="e3c4f-170">[Permissões (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-170">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="e3c4f-171">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-171">X</span></span>    |     <span data-ttu-id="e3c4f-172">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-172">X</span></span>     |    <span data-ttu-id="e3c4f-173">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-173">X</span></span>    |
| <span data-ttu-id="e3c4f-174">[Regra (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-174">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="e3c4f-175">[Regra (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-175">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="e3c4f-176">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-176">X</span></span>    |
| <span data-ttu-id="e3c4f-177">[Requisitos (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-177">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="e3c4f-178">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-178">X</span></span>    |
| <span data-ttu-id="e3c4f-179">[Conjunto\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-179">[Set\*][]</span></span><br/><span data-ttu-id="e3c4f-180">[Conjuntos (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-180">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="e3c4f-181">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-181">X</span></span>    |
| <span data-ttu-id="e3c4f-182">[Formulário\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-182">[Form\*][]</span></span><br/><span data-ttu-id="e3c4f-183">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-183">[formsettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="e3c4f-184">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-184">X</span></span>    |
| <span data-ttu-id="e3c4f-185">[Conjuntos (Requisitos)\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-185">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="e3c4f-186">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-186">X</span></span>    |     <span data-ttu-id="e3c4f-187">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-187">X</span></span>     |         |
| <span data-ttu-id="e3c4f-188">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-188">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="e3c4f-189">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-189">X</span></span>    |     <span data-ttu-id="e3c4f-190">X</span><span class="sxs-lookup"><span data-stu-id="e3c4f-190">X</span></span>     |         |

<span data-ttu-id="e3c4f-191">_\*Adicionados no esquema de manifesto de suplementos da versão 1.1 do Office._</span><span class="sxs-lookup"><span data-stu-id="e3c4f-191">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: https://docs.microsoft.com/javascript/office/manifest/officeapp?view=office-js
[id]: https://docs.microsoft.com/javascript/office/manifest/id
[versão]: https://docs.microsoft.com/javascript/office/manifest/version
[version]: https://docs.microsoft.com/javascript/office/manifest/version
[providername]: https://docs.microsoft.com/javascript/office/manifest/providername
[defaultlocale]: https://docs.microsoft.com/javascript/office/manifest/defaultlocale
[displayname]: https://docs.microsoft.com/javascript/office/manifest/displayname
[descrição]: https://docs.microsoft.com/javascript/office/manifest/description
[description]: https://docs.microsoft.com/javascript/office/manifest/description
[iconurl]: https://docs.microsoft.com/javascript/office/manifest/iconurl
[highresolutioniconurl]: https://docs.microsoft.com/javascript/office/manifest/highresolutioniconurl
[defaultsettings (contentapp)]: https://docs.microsoft.com/javascript/office/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://docs.microsoft.com/javascript/office/manifest/defaultsettings
[sourcelocation (contentapp)]: https://docs.microsoft.com/javascript/office/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://docs.microsoft.com/javascript/office/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissões (contentapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[permissions (contentapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[permissões (taskpaneapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[permissions (taskpaneapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[permissões (mailapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[permissions (mailapp)]: https://docs.microsoft.com/javascript/office/manifest/permissions
[regra (rulecollection)]: https://docs.microsoft.com/javascript/office/manifest/rule
[rule (rulecollection)]: https://docs.microsoft.com/javascript/office/manifest/rule
[regra (mailapp)]: https://docs.microsoft.com/javascript/office/manifest/rule
[rule (mailapp)]: https://docs.microsoft.com/javascript/office/manifest/rule
[requisitos (mailapp)]: https://docs.microsoft.com/javascript/office/manifest/requirements
[requirements (mailapp)\*]: https://docs.microsoft.com/javascript/office/manifest/requirements
[conjunto\*]: https://docs.microsoft.com/javascript/office/manifest/set
[set\*]: https://docs.microsoft.com/javascript/office/manifest/set
[conjuntos (mailapprequirements)\*]: https://docs.microsoft.com/javascript/office/manifest/sets
[sets (mailapprequirements)\*]: https://docs.microsoft.com/javascript/office/manifest/sets
[formulário\*]: https://docs.microsoft.com/javascript/office/manifest/form
[form\*]: https://docs.microsoft.com/javascript/office/manifest/form
[formsettings*]: https://docs.microsoft.com/javascript/office/manifest/formsettings
[conjuntos (requisitos)\*]: https://docs.microsoft.com/javascript/office/manifest/sets
[sets (requirements)\*]: https://docs.microsoft.com/javascript/office/manifest/sets
[hosts*]: https://docs.microsoft.com/javascript/office/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="e3c4f-219">Requisitos de hospedagem</span><span class="sxs-lookup"><span data-stu-id="e3c4f-219">Hosting requirements</span></span>

<span data-ttu-id="e3c4f-220">Todas as imagens de URIs, como as usadas para os [comandos do suplemento][] devem ser compatíveis com armazenamento em cache.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-220">All image URIs, such as those used for [Add-in Commands][], must support caching.</span></span> <span data-ttu-id="e3c4f-221">O servidor que hospeda a imagem não deve retornar um cabeçalho `Cache-Control` especificando `no-cache`, `no-store` ou opções semelhantes na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-221">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="e3c4f-222">Todas as URLs, como os locais dos arquivos de origem especificados no elemento [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation), devem estar **protegidos por SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-222">All URLs, such as the source file locations specified in the [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="e3c4f-223">Práticas recomendadas de envio ao AppSource</span><span class="sxs-lookup"><span data-stu-id="e3c4f-223">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="e3c4f-p103">Verifique se a identificação do suplemento é um GUID válido e exclusivo. Diversas ferramentas de gerador de GUID estão disponíveis na Web e podem ser usadas para criar um GUID exclusivo.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="e3c4f-226">Os suplementos enviados ao AppSource também devem conter o elemento [SupportUrl](https://docs.microsoft.com/javascript/office/manifest/supporturl).</span><span class="sxs-lookup"><span data-stu-id="e3c4f-226">Add-ins submitted to AppSource must also include the [SupportUrl](https://docs.microsoft.com/javascript/office/manifest/supporturl) element.</span></span> <span data-ttu-id="e3c4f-227">Saiba mais em [Políticas de validação para aplicativos e suplementos enviados ao AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="e3c4f-227">For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="e3c4f-228">Use apenas o elemento [AppDomains](https://docs.microsoft.com/javascript/office/manifest/appdomains) para especificar domínios diferentes daqueles especificados no elemento [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation) para cenários de autenticação.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-228">Only use the [AppDomains](https://docs.microsoft.com/javascript/office/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="e3c4f-229">Especificar os domínios que você deseja abrir na janela do suplemento</span><span class="sxs-lookup"><span data-stu-id="e3c4f-229">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="e3c4f-230">Durante a execução no Office Online, o painel de tarefas pode ser navegado para qualquer URL.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-230">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="e3c4f-231">No entanto, em plataformas desktop, se o seu suplemento tentar acessar uma URL em um domínio diferente daquele que hospeda a página inicial (conforme especificado no elemento [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation) do arquivo de manifesto), essa URL será aberta em uma nova janela de navegador fora do painel do suplemento do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-231">By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://docs.microsoft.com/javascript/office/manifest/sourcelocation) element of the manifest file), that URL will open in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="e3c4f-232">Para substituir esse comportamento (do Office para desktop), especifique cada domínio que você deseja abrir na janela do suplemento na lista de domínios especificados no elemento [AppDomains](https://docs.microsoft.com/javascript/office/manifest/appdomains) do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-232">To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://docs.microsoft.com/javascript/office/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="e3c4f-233">Se o suplemento tentar ir para uma URL em um domínio que esteja na lista, ele será aberto no painel de tarefas do Office para desktop e do Office Online.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-233">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="e3c4f-234">Se ele tentar acessar uma URL que não está na lista, no Office para desktop, essa URL será aberta em uma nova janela do navegador (fora do painel do suplemento).</span><span class="sxs-lookup"><span data-stu-id="e3c4f-234">If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="e3c4f-235">Esse comportamento aplica-se somente ao painel raiz do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-235">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="e3c4f-236">Se houver um iframe incorporado na página do suplemento, o iframe poderá ser direcionado para qualquer URL, independentemente de estar listado em **AppDomains**, mesmo no Office para desktop.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-236">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="e3c4f-237">O exemplo de manifesto XML a seguir hospeda a página principal do suplemento no domínio `https://www.contoso.com`, conforme especificado no elemento **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-237">The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="e3c4f-238">Também especifica o domínio `https://www.northwindtraders.com` em um elemento [AppDomain](https://docs.microsoft.com/javascript/office/manifest/appdomain) dentro da lista de elementos **AppDomains**.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-238">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://docs.microsoft.com/javascript/office/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="e3c4f-239">Se o suplemento acessar uma página no domínio www.northwindtraders.com, essa página será aberta no painel do suplemento, mesmo no Office para desktop.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-239">If the add-in goes to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="e3c4f-240">Exemplos e esquemas do arquivo XML do manifesto v1.1</span><span class="sxs-lookup"><span data-stu-id="e3c4f-240">Manifest v1.1 XML file examples and schemas</span></span>
<span data-ttu-id="e3c4f-241">As seções a seguir mostram exemplos de arquivos XML de manifesto v1.1 para suplementos de conteúdo, de painel de tarefas e do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-241">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="e3c4f-242">Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e3c4f-242">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="e3c4f-243">Esquema de manifesto do aplicativo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e3c4f-243">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

<!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

<!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

<!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
   <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
   <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
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
                <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
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
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="e3c4f-244">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e3c4f-244">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="e3c4f-245">Esquema de manifesto do aplicativo de conteúdo</span><span class="sxs-lookup"><span data-stu-id="e3c4f-245">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="e3c4f-246">Email</span><span class="sxs-lookup"><span data-stu-id="e3c4f-246">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="e3c4f-247">Esquema de manifesto do aplicativo de email</span><span class="sxs-lookup"><span data-stu-id="e3c4f-247">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

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
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="e3c4f-248">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="e3c4f-248">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="e3c4f-p109">Para solucionar problemas com seu manifesto, confira [Validar e solucionar problemas com seu manifesto](../testing/troubleshoot-manifest.md). Lá, você encontrará informações sobre como validar o manifesto em relação à [Definição de esquema XML (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) e também como usar o log de tempo de execução para depurar o manifesto.</span><span class="sxs-lookup"><span data-stu-id="e3c4f-p109">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="e3c4f-251">Veja também</span><span class="sxs-lookup"><span data-stu-id="e3c4f-251">See also</span></span>

* <span data-ttu-id="e3c4f-252">[Criar comandos de suplementos em seu manifesto][comandos de suplementos]</span><span class="sxs-lookup"><span data-stu-id="e3c4f-252">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="e3c4f-253">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="e3c4f-253">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="e3c4f-254">Localização para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e3c4f-254">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="e3c4f-255">Referência de esquema para manifestos de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e3c4f-255">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="e3c4f-256">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="e3c4f-256">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[Comandos de suplemento]: create-addin-commands.md
[add-in commands]: create-addin-commands.md