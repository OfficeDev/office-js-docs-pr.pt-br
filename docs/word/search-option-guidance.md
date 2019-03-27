---
title: Usar as opções de pesquisa para localizar o texto no suplemento do Word
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 343271b0863379d799c22f9b63a47a9acfd67b93
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870678"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="85afa-102">Usar as opções de pesquisa para localizar o texto no suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="85afa-102">Use search options to find text in your Word add-in</span></span>

<span data-ttu-id="85afa-103">Os suplementos frequentemente precisam agir com base no texto de um documento.</span><span class="sxs-lookup"><span data-stu-id="85afa-103">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="85afa-104">Uma função de pesquisa é exposta por cada controle de conteúdo (isso inclui [Corpo](/javascript/api/word/word.body), [Parágrafo](/javascript/api/word/word.paragraph), [Intervalo](/javascript/api/word/word.range), [Tabela](/javascript/api/word/word.table), [ Coluna](/javascript/api/word/word.tablerow)e o objeto [ContentControl](/javascript/api/word/word.contentcontrol) base).</span><span class="sxs-lookup"><span data-stu-id="85afa-104">A search function is exposed by every content control (this includes [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), and the base [ContentControl](/javascript/api/word/word.contentcontrol) object).</span></span> <span data-ttu-id="85afa-105">Esta função assume uma cadeia de caracteres (ou expressão curinga) que representa o texto que você está procurando e um objeto [SearchOptions](/javascript/api/word/word.searchoptions).</span><span class="sxs-lookup"><span data-stu-id="85afa-105">This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](/javascript/api/word/word.searchoptions) object.</span></span> <span data-ttu-id="85afa-106">Retorna um conjunto de intervalos que correspondem ao texto de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="85afa-106">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="85afa-107">Opções de pesquisa</span><span class="sxs-lookup"><span data-stu-id="85afa-107">Search options</span></span>

<span data-ttu-id="85afa-108">As opções de pesquisa são uma coleção de valores boolianos que definem como o parâmetro de pesquisa deve ser tratado.</span><span class="sxs-lookup"><span data-stu-id="85afa-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span>

| <span data-ttu-id="85afa-109">Propriedade</span><span class="sxs-lookup"><span data-stu-id="85afa-109">Property</span></span>     | <span data-ttu-id="85afa-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="85afa-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="85afa-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="85afa-111">ignorePunct</span></span>|<span data-ttu-id="85afa-112">Obtém ou define um valor que indica se devem ser ignorados todos os caracteres de pontuação entre as palavras.</span><span class="sxs-lookup"><span data-stu-id="85afa-112">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="85afa-113">Corresponde à caixa de seleção "Ignorar caracteres de pontuação" na caixa de diálogo Localizar e Substituir.</span><span class="sxs-lookup"><span data-stu-id="85afa-113">Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="85afa-114">ignoreSpace</span></span>|<span data-ttu-id="85afa-115">Obtém ou define um valor que indica se devem ser ignorados todos os espaços em branco entre as palavras.</span><span class="sxs-lookup"><span data-stu-id="85afa-115">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="85afa-116">Corresponde à caixa de seleção "Ignorar caracteres de espaço em branco" na caixa de diálogo Localizar e Substituir.</span><span class="sxs-lookup"><span data-stu-id="85afa-116">Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="85afa-117">matchCase</span></span>|<span data-ttu-id="85afa-118">Obtém ou define um valor que determina quando realizar uma pesquisa que diferencia maiúsculas de minúsculas.</span><span class="sxs-lookup"><span data-stu-id="85afa-118">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="85afa-119">Corresponde à caixa de seleção "Diferenciar maiúsculas de minúsculas", na caixa de diálogo "Localizar e Substituir" (menu Editar).</span><span class="sxs-lookup"><span data-stu-id="85afa-119">Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="85afa-120">matchPrefix</span></span>|<span data-ttu-id="85afa-121">Obtém ou define um valor que indica se se deve corresponder palavras que começam com a cadeia de caracteres de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="85afa-121">Gets or sets a value indicating whether to match words that begin with the search string.</span></span> <span data-ttu-id="85afa-122">Corresponde à caixa de seleção "Corresponder prefixo" na caixa de diálogo "Localizar e Substituir".</span><span class="sxs-lookup"><span data-stu-id="85afa-122">Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="85afa-123">matchSuffix</span></span>|<span data-ttu-id="85afa-124">Obtém ou define um valor que indica se se deve corresponder palavras que terminam com a cadeia de caracteres de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="85afa-124">Gets or sets a value indicating whether to match words that end with the search string.</span></span> <span data-ttu-id="85afa-125">Corresponde à caixa de seleção "Corresponder sufixo", na caixa de diálogo "Localizar e Substituir".</span><span class="sxs-lookup"><span data-stu-id="85afa-125">Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="85afa-126">matchWholeWord</span></span>|<span data-ttu-id="85afa-127">Obtém ou define um valor que indica se a operação de localização localiza apenas palavras inteiras, e não texto que faz parte de uma palavra maior.</span><span class="sxs-lookup"><span data-stu-id="85afa-127">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word.</span></span> <span data-ttu-id="85afa-128">Corresponde à caixa de seleção "Localizar apenas palavras inteiras" na caixa de diálogo Localizar e Substituir.</span><span class="sxs-lookup"><span data-stu-id="85afa-128">Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="85afa-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="85afa-129">matchWildcards</span></span>|<span data-ttu-id="85afa-130">Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais.</span><span class="sxs-lookup"><span data-stu-id="85afa-130">Gets or sets a value indicating whether the search will be performed using special search operators.</span></span> <span data-ttu-id="85afa-131">Corresponde à caixa de seleção "Usar caracteres curinga" na caixa de diálogo "Localizar e Substituir".</span><span class="sxs-lookup"><span data-stu-id="85afa-131">Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="85afa-132">Diretrizes para caracteres curinga</span><span class="sxs-lookup"><span data-stu-id="85afa-132">Wildcard guidance</span></span>

<span data-ttu-id="85afa-133">A tabela a seguir fornece orientações em torno de caracteres curinga de pesquisa da API JavaScript do Word.</span><span class="sxs-lookup"><span data-stu-id="85afa-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="85afa-134">Para localizar:</span><span class="sxs-lookup"><span data-stu-id="85afa-134">To find:</span></span>         | <span data-ttu-id="85afa-135">Curinga</span><span class="sxs-lookup"><span data-stu-id="85afa-135">Wildcard</span></span> |  <span data-ttu-id="85afa-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="85afa-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="85afa-137">Qualquer caractere simples</span><span class="sxs-lookup"><span data-stu-id="85afa-137">Any single character</span></span>| <span data-ttu-id="85afa-138">?</span><span class="sxs-lookup"><span data-stu-id="85afa-138"></span></span> |<span data-ttu-id="85afa-139">c?l localiza "calor" e "caldo".</span><span class="sxs-lookup"><span data-stu-id="85afa-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="85afa-140">Qualquer cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="85afa-140">Any string of characters</span></span>| * |<span data-ttu-id="85afa-141">g\*s localiza gostar e gastar.</span><span class="sxs-lookup"><span data-stu-id="85afa-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="85afa-142">O início de uma palavra</span><span class="sxs-lookup"><span data-stu-id="85afa-142">The beginning of a word</span></span>|< |<span data-ttu-id="85afa-143">< (inter) localiza interseção e interessante, mas não localiza desinteresse.</span><span class="sxs-lookup"><span data-stu-id="85afa-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="85afa-144">O final de uma palavra</span><span class="sxs-lookup"><span data-stu-id="85afa-144">The end of a word</span></span> |> |<span data-ttu-id="85afa-145">(em)> localiza vargem e miragem, mas não localiza embrião.</span><span class="sxs-lookup"><span data-stu-id="85afa-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="85afa-146">Um dos caracteres especificados</span><span class="sxs-lookup"><span data-stu-id="85afa-146">One of the specified characters</span></span>|<span data-ttu-id="85afa-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="85afa-147"></span></span> |<span data-ttu-id="85afa-148">t[eo]m localiza tem e tom.</span><span class="sxs-lookup"><span data-stu-id="85afa-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="85afa-149">Qualquer caractere único deste intervalo</span><span class="sxs-lookup"><span data-stu-id="85afa-149">Any single character in this range</span></span>| <span data-ttu-id="85afa-150">[-]</span><span class="sxs-lookup"><span data-stu-id="85afa-150"></span></span> |<span data-ttu-id="85afa-p109">[r-t]olo localiza rolo e solo. Os intervalos devem estar em ordem crescente.</span><span class="sxs-lookup"><span data-stu-id="85afa-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="85afa-153">Qualquer caractere único, exceto os caracteres do intervalo entre colchetes</span><span class="sxs-lookup"><span data-stu-id="85afa-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="85afa-155">t[!a-m]que localiza toque e trunque, mas não localiza taque ou tique.</span><span class="sxs-lookup"><span data-stu-id="85afa-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="85afa-156">Número de ocorrências exatas do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="85afa-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="85afa-157">{n}</span><span class="sxs-lookup"><span data-stu-id="85afa-157">{n}</span></span> |<span data-ttu-id="85afa-158">ve{2}m localiza veem, mas não vem.</span><span class="sxs-lookup"><span data-stu-id="85afa-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="85afa-159">Número mínimo de ocorrências do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="85afa-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="85afa-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="85afa-160">{n,}</span></span> |<span data-ttu-id="85afa-161">ve{1,}m localiza vem e veem.</span><span class="sxs-lookup"><span data-stu-id="85afa-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="85afa-162">Número de ocorrências do caractere ou expressão anterior dentro de um intervalo</span><span class="sxs-lookup"><span data-stu-id="85afa-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="85afa-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="85afa-163">{n,m}</span></span> |<span data-ttu-id="85afa-164">10{1,3} localiza 10, 100 e 1000.</span><span class="sxs-lookup"><span data-stu-id="85afa-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="85afa-165">Uma ou mais ocorrências do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="85afa-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="85afa-166">re@r localiza reter e reverter.</span><span class="sxs-lookup"><span data-stu-id="85afa-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="85afa-167">Escapar os caracteres especiais</span><span class="sxs-lookup"><span data-stu-id="85afa-167">Escaping the special characters</span></span>

<span data-ttu-id="85afa-p110">A pesquisa com caracteres curinga é essencialmente igual à pesquisa em uma expressão regular. Há caracteres especiais em expressões regulares, como “[', ']”, “(', ')”, “{”, “}”, “\*”, “?”, “<”, “>”, “!” e “@”. Se um desses caracteres fizer parte da cadeia de caracteres literal que o código está procurando, ele precisará ser escapado para que o Word saiba que ele deve ser tratado literalmente e não como parte da lógica da expressão regular. Para escapar um caractere na pesquisa da interface de usuário do Word, prefixe-o com um caractere “\'”, mas, para escapá-lo programaticamente, coloque-o entre caracteres “[]”. Por exemplo, “[\*]\*” pesquisa qualquer cadeia de caracteres que comece com “\*” seguido por qualquer número de outros caracteres.</span><span class="sxs-lookup"><span data-stu-id="85afa-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="85afa-173">Exemplos</span><span class="sxs-lookup"><span data-stu-id="85afa-173">Examples</span></span>

<span data-ttu-id="85afa-174">Os exemplos a seguir demonstram cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="85afa-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="85afa-175">Ignorar pesquisa de pontuação</span><span class="sxs-lookup"><span data-stu-id="85afa-175">Ignore punctuation search</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="85afa-176">Pesquisa com base em um prefixo</span><span class="sxs-lookup"><span data-stu-id="85afa-176">Search based on a prefix</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="85afa-177">Pesquisa com base em um sufixo</span><span class="sxs-lookup"><span data-stu-id="85afa-177">Search based on a suffix</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a><span data-ttu-id="85afa-178">Pesquisa usando caracteres curinga</span><span class="sxs-lookup"><span data-stu-id="85afa-178">Search using a wildcard</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="85afa-179">Mais informações podem ser encontradas na [Referência de API JavaScript do Word](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="85afa-179">More information can be found in the [Word JavaScript Reference API](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview).</span></span>
