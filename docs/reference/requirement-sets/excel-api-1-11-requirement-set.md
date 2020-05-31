---
title: Conjunto de requisitos de API JavaScript do Excel 1,11
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,11
ms.date: 05/06/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 13ec5e944a2b7ea05b1054939c2c4c6438f26009
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413228"
---
# <a name="whats-new-in-excel-javascript-api-111"></a><span data-ttu-id="6c0b5-103">O que há de novo na API JavaScript do Excel 1,11</span><span class="sxs-lookup"><span data-stu-id="6c0b5-103">What's new in Excel JavaScript API 1.11</span></span>

<span data-ttu-id="6c0b5-104">O ExcelApi 1,11 melhorou o suporte para comentários e controles de nível de pasta de trabalho (como salvar e fechar a pasta de trabalho).</span><span class="sxs-lookup"><span data-stu-id="6c0b5-104">The ExcelApi 1.11 improved support for comments and workbook-level controls (such as saving and closing the workbook).</span></span> <span data-ttu-id="6c0b5-105">Ele também adicionou acesso às configurações de cultura para ajudar a sua conta na localização.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-105">It also added access to culture settings to help account for localization.</span></span>

| <span data-ttu-id="6c0b5-106">Área de recurso</span><span class="sxs-lookup"><span data-stu-id="6c0b5-106">Feature area</span></span> | <span data-ttu-id="6c0b5-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="6c0b5-107">Description</span></span> | <span data-ttu-id="6c0b5-108">Objetos relevantes</span><span class="sxs-lookup"><span data-stu-id="6c0b5-108">Relevant objects</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="6c0b5-109">Comentários [menciona](../../excel/excel-add-ins-comments.md#mentions)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-109">Comment [Mentions](../../excel/excel-add-ins-comments.md#mentions)</span></span> |<span data-ttu-id="6c0b5-110">Marca e notifica outros usuários da pasta de trabalho por meio de comentários.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-110">Tags and notifies other workbook users through comments.</span></span> | <span data-ttu-id="6c0b5-111">[Comentário](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-111">[Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent)</span></span> |
| <span data-ttu-id="6c0b5-112">[Resolução](../../excel/excel-add-ins-comments.md#resolve-comment-threads) de comentários</span><span class="sxs-lookup"><span data-stu-id="6c0b5-112">Comment [Resolution](../../excel/excel-add-ins-comments.md#resolve-comment-threads)</span></span> | <span data-ttu-id="6c0b5-113">Resolver os threads de comentário e obter o status de resolução.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-113">Resolve comment threads and get the resolution status.</span></span> | [<span data-ttu-id="6c0b5-114">Comment</span><span class="sxs-lookup"><span data-stu-id="6c0b5-114">Comment</span></span>](/javascript/api/excel/excel.comment) |
| [<span data-ttu-id="6c0b5-115">Configurações de cultura</span><span class="sxs-lookup"><span data-stu-id="6c0b5-115">Culture settings</span></span>](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | <span data-ttu-id="6c0b5-116">Obtém configurações culturais do sistema para a pasta de trabalho, como formatação de número.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-116">Gets cultural system settings for the workbook, such as number formatting.</span></span> | <span data-ttu-id="6c0b5-117">[CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-117">[CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application)</span></span> |
| [<span data-ttu-id="6c0b5-118">Recortar e colar (moveTo)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-118">Cut and paste (moveTo)</span></span>](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | <span data-ttu-id="6c0b5-119">Replica a funcionalidade de recortar e colar no Excel para um intervalo.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-119">Replicates the cut-and-paste functionality in Excel for a Range.</span></span> | [<span data-ttu-id="6c0b5-120">Range</span><span class="sxs-lookup"><span data-stu-id="6c0b5-120">Range</span></span>](/javascript/api/excel/excel.range) |
| <span data-ttu-id="6c0b5-121">[Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook) a pasta de trabalho</span><span class="sxs-lookup"><span data-stu-id="6c0b5-121">Workbook [Save](../../excel/excel-add-ins-workbooks.md#save-the-workbook) and [Close](../../excel/excel-add-ins-workbooks.md#close-the-workbook)</span></span> | <span data-ttu-id="6c0b5-122">Salve e feche a pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-122">Save and close workbooks.</span></span> | [<span data-ttu-id="6c0b5-123">Workbook</span><span class="sxs-lookup"><span data-stu-id="6c0b5-123">Workbook</span></span>](/javascript/api/excel/excel.workbook) |
| <span data-ttu-id="6c0b5-124">Eventos de planilha</span><span class="sxs-lookup"><span data-stu-id="6c0b5-124">Worksheet events</span></span> | <span data-ttu-id="6c0b5-125">Eventos adicionais e informações de eventos para cálculos de planilha e linhas ocultas.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-125">Additional events and event information for worksheet calculations and hidden rows.</span></span> | <span data-ttu-id="6c0b5-126">[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-126">[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)</span></span> |

## <a name="api-list"></a><span data-ttu-id="6c0b5-127">Lista de APIs</span><span class="sxs-lookup"><span data-stu-id="6c0b5-127">API list</span></span>

<span data-ttu-id="6c0b5-128">A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,11.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-128">The following table lists the APIs in Excel JavaScript API requirement set 1.11.</span></span> <span data-ttu-id="6c0b5-129">Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,11 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,11 ou anterior](/javascript/api/excel?view=excel-js-1.11).</span><span class="sxs-lookup"><span data-stu-id="6c0b5-129">To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.11 or earlier, see [Excel APIs in requirement set 1.11 or earlier](/javascript/api/excel?view=excel-js-1.11).</span></span>

| <span data-ttu-id="6c0b5-130">Classe</span><span class="sxs-lookup"><span data-stu-id="6c0b5-130">Class</span></span> | <span data-ttu-id="6c0b5-131">Campos</span><span class="sxs-lookup"><span data-stu-id="6c0b5-131">Fields</span></span> | <span data-ttu-id="6c0b5-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="6c0b5-132">Description</span></span> |
|:---|:---|:---|
|[<span data-ttu-id="6c0b5-133">Aplicativo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-133">Application</span></span>](/javascript/api/excel/excel.application)|[<span data-ttu-id="6c0b5-134">cultureInfo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-134">cultureInfo</span></span>](/javascript/api/excel/excel.application#cultureinfo)|<span data-ttu-id="6c0b5-135">Fornece informações com base nas configurações de cultura do sistema atual.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-135">Provides information based on current system culture settings.</span></span> <span data-ttu-id="6c0b5-136">Isso inclui os nomes de cultura, a formatação de números e outras configurações dependentes de cultura.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-136">This includes the culture names, number formatting, and other culturally dependent settings.</span></span>|
||[<span data-ttu-id="6c0b5-137">decimalSeparator</span><span class="sxs-lookup"><span data-stu-id="6c0b5-137">decimalSeparator</span></span>](/javascript/api/excel/excel.application#decimalseparator)|<span data-ttu-id="6c0b5-138">Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-138">Gets the string used as the decimal separator for numeric values.</span></span> <span data-ttu-id="6c0b5-139">Isso é baseado nas configurações locais do Excel.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-139">This is based on Excel's local settings.</span></span>|
||[<span data-ttu-id="6c0b5-140">thousandsSeparator</span><span class="sxs-lookup"><span data-stu-id="6c0b5-140">thousandsSeparator</span></span>](/javascript/api/excel/excel.application#thousandsseparator)|<span data-ttu-id="6c0b5-141">Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-141">Gets the string used to separate groups of digits to the left of the decimal for numeric values.</span></span> <span data-ttu-id="6c0b5-142">Isso é baseado nas configurações locais do Excel.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-142">This is based on Excel's local settings.</span></span>|
||[<span data-ttu-id="6c0b5-143">useSystemSeparators</span><span class="sxs-lookup"><span data-stu-id="6c0b5-143">useSystemSeparators</span></span>](/javascript/api/excel/excel.application#usesystemseparators)|<span data-ttu-id="6c0b5-144">Especifica se os separadores de sistema do Excel estão habilitados.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-144">Specifies if the system separators of Excel are enabled.</span></span>|
|[<span data-ttu-id="6c0b5-145">Comment</span><span class="sxs-lookup"><span data-stu-id="6c0b5-145">Comment</span></span>](/javascript/api/excel/excel.comment)|[<span data-ttu-id="6c0b5-146">menções</span><span class="sxs-lookup"><span data-stu-id="6c0b5-146">mentions</span></span>](/javascript/api/excel/excel.comment#mentions)|<span data-ttu-id="6c0b5-147">Obtém as entidades (por exemplo, pessoas) mencionadas em comentários.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-147">Gets the entities (e.g., people) that are mentioned in comments.</span></span>|
||[<span data-ttu-id="6c0b5-148">richContent</span><span class="sxs-lookup"><span data-stu-id="6c0b5-148">richContent</span></span>](/javascript/api/excel/excel.comment#richcontent)|<span data-ttu-id="6c0b5-149">Obtém o conteúdo de comentário avançado (por exemplo, menciona em comentários).</span><span class="sxs-lookup"><span data-stu-id="6c0b5-149">Gets the rich comment content (e.g., mentions in comments).</span></span> <span data-ttu-id="6c0b5-150">Essa cadeia de caracteres não deve ser exibida para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-150">This string is not meant to be displayed to end-users.</span></span> <span data-ttu-id="6c0b5-151">Seu suplemento só deve usar este para analisar conteúdo de comentário avançado.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-151">Your add-in should only use this to parse rich comment content.</span></span>|
||[<span data-ttu-id="6c0b5-152">Obtido</span><span class="sxs-lookup"><span data-stu-id="6c0b5-152">resolved</span></span>](/javascript/api/excel/excel.comment#resolved)|<span data-ttu-id="6c0b5-153">O status do thread de comentários.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-153">The comment thread status.</span></span> <span data-ttu-id="6c0b5-154">O valor "true" significa que o thread de comentários é resolvido.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-154">A value of "true" means the comment thread is resolved.</span></span>|
||[<span data-ttu-id="6c0b5-155">updateMentions (contentWithMentions: Excel. CommentRichContent)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-155">updateMentions(contentWithMentions: Excel.CommentRichContent)</span></span>](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|<span data-ttu-id="6c0b5-156">Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-156">Updates the comment content with a specially formatted string and a list of mentions.</span></span>|
|[<span data-ttu-id="6c0b5-157">CommentCollection</span><span class="sxs-lookup"><span data-stu-id="6c0b5-157">CommentCollection</span></span>](/javascript/api/excel/excel.commentcollection)|[<span data-ttu-id="6c0b5-158">Add (cellAddress: \| String de intervalo, Content: \| cadeia de caracteres CommentRichContent, ContentType?: Excel. ContentType)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-158">add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)</span></span>](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|<span data-ttu-id="6c0b5-159">Cria um novo comentário com o conteúdo fornecido na célula especificada.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-159">Creates a new comment with the given content on the given cell.</span></span> <span data-ttu-id="6c0b5-160">Um `InvalidArgument` erro será acionado se o intervalo fornecido for maior que uma célula.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-160">An `InvalidArgument` error is thrown if the provided range is larger than one cell.</span></span>|
|[<span data-ttu-id="6c0b5-161">CommentMention</span><span class="sxs-lookup"><span data-stu-id="6c0b5-161">CommentMention</span></span>](/javascript/api/excel/excel.commentmention)|[<span data-ttu-id="6c0b5-162">email</span><span class="sxs-lookup"><span data-stu-id="6c0b5-162">email</span></span>](/javascript/api/excel/excel.commentmention#email)|<span data-ttu-id="6c0b5-163">O endereço de email da entidade mencionada em comentário.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-163">The email address of the entity that is mentioned in comment.</span></span>|
||[<span data-ttu-id="6c0b5-164">id</span><span class="sxs-lookup"><span data-stu-id="6c0b5-164">id</span></span>](/javascript/api/excel/excel.commentmention#id)|<span data-ttu-id="6c0b5-165">A ID da entidade.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-165">The id of the entity.</span></span> <span data-ttu-id="6c0b5-166">A ID corresponde a uma das IDs no `CommentRichContent.richContent` .</span><span class="sxs-lookup"><span data-stu-id="6c0b5-166">The id matches one of the ids in `CommentRichContent.richContent`.</span></span>|
||[<span data-ttu-id="6c0b5-167">name</span><span class="sxs-lookup"><span data-stu-id="6c0b5-167">name</span></span>](/javascript/api/excel/excel.commentmention#name)|<span data-ttu-id="6c0b5-168">O nome da entidade mencionada em comentário.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-168">The name of the entity that is mentioned in comment.</span></span>|
|[<span data-ttu-id="6c0b5-169">CommentReply</span><span class="sxs-lookup"><span data-stu-id="6c0b5-169">CommentReply</span></span>](/javascript/api/excel/excel.commentreply)|[<span data-ttu-id="6c0b5-170">menções</span><span class="sxs-lookup"><span data-stu-id="6c0b5-170">mentions</span></span>](/javascript/api/excel/excel.commentreply#mentions)|<span data-ttu-id="6c0b5-171">As entidades (por exemplo, pessoas) mencionadas em comentários.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-171">The entities (e.g., people) that are mentioned in comments.</span></span>|
||[<span data-ttu-id="6c0b5-172">Obtido</span><span class="sxs-lookup"><span data-stu-id="6c0b5-172">resolved</span></span>](/javascript/api/excel/excel.commentreply#resolved)|<span data-ttu-id="6c0b5-173">O status de resposta de comentário.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-173">The comment reply status.</span></span> <span data-ttu-id="6c0b5-174">O valor "true" significa que a resposta está no estado resolvido.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-174">A value of "true" means the reply is in the resolved state.</span></span>|
||[<span data-ttu-id="6c0b5-175">richContent</span><span class="sxs-lookup"><span data-stu-id="6c0b5-175">richContent</span></span>](/javascript/api/excel/excel.commentreply#richcontent)|<span data-ttu-id="6c0b5-176">O conteúdo de comentário avançado (por exemplo, menciona comentários).</span><span class="sxs-lookup"><span data-stu-id="6c0b5-176">The rich comment content (e.g., mentions in comments).</span></span> <span data-ttu-id="6c0b5-177">Essa cadeia de caracteres não deve ser exibida para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-177">This string is not meant to be displayed to end-users.</span></span> <span data-ttu-id="6c0b5-178">Seu suplemento só deve usar este para analisar conteúdo de comentário avançado.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-178">Your add-in should only use this to parse rich comment content.</span></span>|
||[<span data-ttu-id="6c0b5-179">updateMentions (contentWithMentions: Excel. CommentRichContent)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-179">updateMentions(contentWithMentions: Excel.CommentRichContent)</span></span>](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|<span data-ttu-id="6c0b5-180">Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-180">Updates the comment content with a specially formatted string and a list of mentions.</span></span>|
|[<span data-ttu-id="6c0b5-181">CommentReplyCollection</span><span class="sxs-lookup"><span data-stu-id="6c0b5-181">CommentReplyCollection</span></span>](/javascript/api/excel/excel.commentreplycollection)|[<span data-ttu-id="6c0b5-182">Add (Content: CommentRichContent \| String, ContentType?: Excel. ContentType)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-182">add(content: CommentRichContent \| string, contentType?: Excel.ContentType)</span></span>](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|<span data-ttu-id="6c0b5-183">Cria uma resposta de comentário para o comentário.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-183">Creates a comment reply for comment.</span></span>|
|[<span data-ttu-id="6c0b5-184">CommentRichContent</span><span class="sxs-lookup"><span data-stu-id="6c0b5-184">CommentRichContent</span></span>](/javascript/api/excel/excel.commentrichcontent)|[<span data-ttu-id="6c0b5-185">menções</span><span class="sxs-lookup"><span data-stu-id="6c0b5-185">mentions</span></span>](/javascript/api/excel/excel.commentrichcontent#mentions)|<span data-ttu-id="6c0b5-186">Uma matriz que contém todas as entidades (por exemplo, pessoas) mencionadas no comentário.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-186">An array containing all the entities (e.g., people) mentioned within the comment.</span></span>|
||[<span data-ttu-id="6c0b5-187">richContent</span><span class="sxs-lookup"><span data-stu-id="6c0b5-187">richContent</span></span>](/javascript/api/excel/excel.commentrichcontent#richcontent)|<span data-ttu-id="6c0b5-188">Especifica o conteúdo avançado do comentário (por exemplo, conteúdo de comentários com menção, a primeira entidade mencionada tem um atributo ID 0 e a segunda entidade mencionada tem um atributo ID de 1.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-188">Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an id attribute of 0, and the second mentioned entity has an id attribute of 1.</span></span>|
|[<span data-ttu-id="6c0b5-189">CultureInfo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-189">CultureInfo</span></span>](/javascript/api/excel/excel.cultureinfo)|[<span data-ttu-id="6c0b5-190">name</span><span class="sxs-lookup"><span data-stu-id="6c0b5-190">name</span></span>](/javascript/api/excel/excel.cultureinfo#name)|<span data-ttu-id="6c0b5-191">Obtém o nome da cultura no formato languagecode2-Country/regioncode2 (por exemplo, "zh-CN" ou "en-US").</span><span class="sxs-lookup"><span data-stu-id="6c0b5-191">Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us").</span></span> <span data-ttu-id="6c0b5-192">Isso é baseado nas configurações atuais do sistema.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-192">This is based on current system settings.</span></span>|
||[<span data-ttu-id="6c0b5-193">numberFormat</span><span class="sxs-lookup"><span data-stu-id="6c0b5-193">numberFormat</span></span>](/javascript/api/excel/excel.cultureinfo#numberformat)|<span data-ttu-id="6c0b5-194">Define o formato culturalmente apropriado para exibir números.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-194">Defines the culturally appropriate format of displaying numbers.</span></span> <span data-ttu-id="6c0b5-195">Isso é baseado nas configurações atuais de cultura do sistema.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-195">This is based on current system culture settings.</span></span>|
|[<span data-ttu-id="6c0b5-196">NumberFormatInfo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-196">NumberFormatInfo</span></span>](/javascript/api/excel/excel.numberformatinfo)|[<span data-ttu-id="6c0b5-197">numberDecimalSeparator</span><span class="sxs-lookup"><span data-stu-id="6c0b5-197">numberDecimalSeparator</span></span>](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|<span data-ttu-id="6c0b5-198">Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-198">Gets the string used as the decimal separator for numeric values.</span></span> <span data-ttu-id="6c0b5-199">Isso é baseado nas configurações atuais do sistema.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-199">This is based on current system settings.</span></span>|
||[<span data-ttu-id="6c0b5-200">numberGroupSeparator</span><span class="sxs-lookup"><span data-stu-id="6c0b5-200">numberGroupSeparator</span></span>](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|<span data-ttu-id="6c0b5-201">Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-201">Gets the string used to separate groups of digits to the left of the decimal for numeric values.</span></span> <span data-ttu-id="6c0b5-202">Isso é baseado nas configurações atuais do sistema.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-202">This is based on current system settings.</span></span>|
|[<span data-ttu-id="6c0b5-203">Range</span><span class="sxs-lookup"><span data-stu-id="6c0b5-203">Range</span></span>](/javascript/api/excel/excel.range)|[<span data-ttu-id="6c0b5-204">moveTo (destinationRange: cadeia de caracteres de intervalo \| )</span><span class="sxs-lookup"><span data-stu-id="6c0b5-204">moveTo(destinationRange: Range \| string)</span></span>](/javascript/api/excel/excel.range#moveto-destinationrange-)|<span data-ttu-id="6c0b5-205">Move valores de célula, formatação e fórmulas do intervalo atual para o intervalo de destino, substituindo as informações antigas nessas células.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-205">Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.</span></span>|
|[<span data-ttu-id="6c0b5-206">RangeFormat</span><span class="sxs-lookup"><span data-stu-id="6c0b5-206">RangeFormat</span></span>](/javascript/api/excel/excel.rangeformat)|[<span data-ttu-id="6c0b5-207">adjustIndent (valor: número)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-207">adjustIndent(amount: number)</span></span>](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|<span data-ttu-id="6c0b5-208">Ajusta o recuo da formatação do intervalo.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-208">Adjusts the indentation of the range formatting.</span></span> <span data-ttu-id="6c0b5-209">O valor de recuo varia de 0 a 250 e é medido em caracteres.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-209">The indent value ranges from 0 to 250 and is measured in characters.</span></span>|
|[<span data-ttu-id="6c0b5-210">Workbook</span><span class="sxs-lookup"><span data-stu-id="6c0b5-210">Workbook</span></span>](/javascript/api/excel/excel.workbook)|[<span data-ttu-id="6c0b5-211">close(closeBehavior?: Excel.CloseBehavior)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-211">close(closeBehavior?: Excel.CloseBehavior)</span></span>](/javascript/api/excel/excel.workbook#close-closebehavior-)|<span data-ttu-id="6c0b5-212">Fechar a pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-212">Close current workbook.</span></span>|
||[<span data-ttu-id="6c0b5-213">save(saveBehavior?: Excel.SaveBehavior)</span><span class="sxs-lookup"><span data-stu-id="6c0b5-213">save(saveBehavior?: Excel.SaveBehavior)</span></span>](/javascript/api/excel/excel.workbook#save-savebehavior-)|<span data-ttu-id="6c0b5-214">Salvar a pasta de trabalho atual.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-214">Save current workbook.</span></span>|
|[<span data-ttu-id="6c0b5-215">Worksheet</span><span class="sxs-lookup"><span data-stu-id="6c0b5-215">Worksheet</span></span>](/javascript/api/excel/excel.worksheet)|[<span data-ttu-id="6c0b5-216">onRowHiddenChanged</span><span class="sxs-lookup"><span data-stu-id="6c0b5-216">onRowHiddenChanged</span></span>](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|<span data-ttu-id="6c0b5-217">Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-217">Occurs when the hidden state of one or more rows has changed on a specific worksheet.</span></span>|
|[<span data-ttu-id="6c0b5-218">WorksheetCalculatedEventArgs</span><span class="sxs-lookup"><span data-stu-id="6c0b5-218">WorksheetCalculatedEventArgs</span></span>](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[<span data-ttu-id="6c0b5-219">address</span><span class="sxs-lookup"><span data-stu-id="6c0b5-219">address</span></span>](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|<span data-ttu-id="6c0b5-220">O endereço do intervalo que concluiu o cálculo.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-220">The address of the range that completed calculation.</span></span>|
|[<span data-ttu-id="6c0b5-221">WorksheetCollection</span><span class="sxs-lookup"><span data-stu-id="6c0b5-221">WorksheetCollection</span></span>](/javascript/api/excel/excel.worksheetcollection)|[<span data-ttu-id="6c0b5-222">onRowHiddenChanged</span><span class="sxs-lookup"><span data-stu-id="6c0b5-222">onRowHiddenChanged</span></span>](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|<span data-ttu-id="6c0b5-223">Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-223">Occurs when the hidden state of one or more rows has changed on a specific worksheet.</span></span>|
|[<span data-ttu-id="6c0b5-224">WorksheetRowHiddenChangedEventArgs</span><span class="sxs-lookup"><span data-stu-id="6c0b5-224">WorksheetRowHiddenChangedEventArgs</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[<span data-ttu-id="6c0b5-225">address</span><span class="sxs-lookup"><span data-stu-id="6c0b5-225">address</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|<span data-ttu-id="6c0b5-226">Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-226">Gets the range address that represents the changed area of a specific worksheet.</span></span>|
||[<span data-ttu-id="6c0b5-227">changeType</span><span class="sxs-lookup"><span data-stu-id="6c0b5-227">changeType</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|<span data-ttu-id="6c0b5-228">Obtém o tipo de alteração que representa como o evento foi acionado.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-228">Gets the type of change that represents how the event was triggered.</span></span> <span data-ttu-id="6c0b5-229">Confira `Excel.RowHiddenChangeType` para obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-229">See `Excel.RowHiddenChangeType` for details.</span></span>|
||[<span data-ttu-id="6c0b5-230">source</span><span class="sxs-lookup"><span data-stu-id="6c0b5-230">source</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|<span data-ttu-id="6c0b5-231">Obtém a origem do evento.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-231">Gets the source of the event.</span></span> <span data-ttu-id="6c0b5-232">Para saber detalhes, confira Excel.EventSource.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-232">See Excel.EventSource for details.</span></span>|
||[<span data-ttu-id="6c0b5-233">tipo</span><span class="sxs-lookup"><span data-stu-id="6c0b5-233">type</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|<span data-ttu-id="6c0b5-234">Obtém o tipo do evento.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-234">Gets the type of the event.</span></span> <span data-ttu-id="6c0b5-235">Para saber detalhes, confira Excel.EventType.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-235">See Excel.EventType for details.</span></span>|
||[<span data-ttu-id="6c0b5-236">worksheetId</span><span class="sxs-lookup"><span data-stu-id="6c0b5-236">worksheetId</span></span>](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|<span data-ttu-id="6c0b5-237">Obtém o id da planilha na qual os dados são alterados.</span><span class="sxs-lookup"><span data-stu-id="6c0b5-237">Gets the id of the worksheet in which the data changed.</span></span>|

## <a name="see-also"></a><span data-ttu-id="6c0b5-238">Confira também</span><span class="sxs-lookup"><span data-stu-id="6c0b5-238">See also</span></span>

- [<span data-ttu-id="6c0b5-239">Documentação deReferência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="6c0b5-239">Excel JavaScript API Reference Documentation</span></span>](/javascript/api/excel?view=excel-js-1.11)
- [<span data-ttu-id="6c0b5-240">Conjuntos de requisitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="6c0b5-240">Excel JavaScript API requirement sets</span></span>](excel-api-requirement-sets.md)