---
title: Usar as opções de pesquisa para encontrar texto no seu suplemento do Word
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: ca5c819edb7f3c183379d9df997e41eb56a4de51
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505367"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="b2233-102">Usar as opções de pesquisa para encontrar texto no seu suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="b2233-102">Use search options to find text in your Word add-in</span></span> 

<span data-ttu-id="b2233-p101">Os suplementos frequentemente precisam agir com base no texto de um documento. Uma função de pesquisa é exposta por todo controle de conteúdo (Isso inclui [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js) e o objeto base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js)). Essa função usa uma sequência de caracteres (ou expressão curinga) que representa o texto que você está procurando e um objeto [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js). Ela retorna uma coleção de intervalos que correspondem ao texto da pesquisa.</span><span class="sxs-lookup"><span data-stu-id="b2233-p101">Add-ins frequently need to act based on the text of a document. A search function is exposed by every content control (this includes [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js), and the base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) object). This function takes in a string (or wldcard expression) representing the text you are searching for and a [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) object. It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="b2233-107">Opções de pesquisa</span><span class="sxs-lookup"><span data-stu-id="b2233-107">Search options</span></span>
<span data-ttu-id="b2233-108">As opções de pesquisa são uma coleção de valores booleanos que definem como o parâmetro de pesquisa deve ser tratado.</span><span class="sxs-lookup"><span data-stu-id="b2233-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span> 

| <span data-ttu-id="b2233-109">Propriedade</span><span class="sxs-lookup"><span data-stu-id="b2233-109">Property</span></span>     | <span data-ttu-id="b2233-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="b2233-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="b2233-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="b2233-111">ignorePunct</span></span>|<span data-ttu-id="b2233-p102">Obtém ou define um valor indicando se deve ignorar todos os caracteres de pontuação entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de pontuação", na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p102">Gets or sets a value indicating whether to ignore all punctuation characters between words. Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="b2233-114">ignoreSpace</span></span>|<span data-ttu-id="b2233-p103">Obtém ou define um valor indicando se deve ignorar todos os espaços em branco entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de espaço em branco", na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p103">Gets or sets a value indicating whether to ignore all whitespace between words. Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="b2233-117">matchCase</span></span>|<span data-ttu-id="b2233-p104">Obtém ou define um valor indicando se deve ser realizada uma pesquisa que diferencia maiúsculas de minúsculas. Corresponde à caixa de seleção "Diferenciar maiúsculas/minúsculas" na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p104">Gets or sets a value indicating whether to perform a case sensitive search. Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="b2233-120">matchPrefix</span></span>|<span data-ttu-id="b2233-p105">Obtém ou define um valor que determina quando fazer correspondência com as palavras que começam com a sequência de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir prefixo", na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p105">Gets or sets a value indicating whether to match words that begin with the search string. Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="b2233-123">matchSuffix</span></span>|<span data-ttu-id="b2233-p106">Obtém ou define um valor que determina quando fazer correspondência com as palavras que terminam com a sequência de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir sufixo", na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p106">Gets or sets a value indicating whether to match words that end with the search string. Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="b2233-126">matchWholeWord</span></span>|<span data-ttu-id="b2233-p107">Obtém ou define um valor indicando a operação para localizar somente palavras inteiras, não o texto que faz parte de uma palavra maior. Corresponde à caixa de seleção "Localizar apenas palavras inteiras", na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p107">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word. Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="b2233-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="b2233-129">matchWildcards</span></span>|<span data-ttu-id="b2233-p108">Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais. Corresponde à caixa de seleção "Usar caracteres curinga" na caixa de diálogo "Localizar e substituir".</span><span class="sxs-lookup"><span data-stu-id="b2233-p108">Gets or sets a value indicating whether the search will be performed using special search operators. Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="b2233-132">Diretrizes para caracteres curinga</span><span class="sxs-lookup"><span data-stu-id="b2233-132">Wildcard Guidance</span></span>
<span data-ttu-id="b2233-133">A tabela a seguir fornece diretrizes sobre os curingas de pesquisa da API JavaScript do Word.</span><span class="sxs-lookup"><span data-stu-id="b2233-133">The following table provides guidance around the Word JavaScript API’s search wildcards.</span></span>

| <span data-ttu-id="b2233-134">Para localizar:</span><span class="sxs-lookup"><span data-stu-id="b2233-134">To find:</span></span>         | <span data-ttu-id="b2233-135">Curinga</span><span class="sxs-lookup"><span data-stu-id="b2233-135">Wildcard</span></span> |  <span data-ttu-id="b2233-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b2233-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="b2233-137">Qualquer caractere simples</span><span class="sxs-lookup"><span data-stu-id="b2233-137">Any single character</span></span>| <span data-ttu-id="b2233-138">?</span><span class="sxs-lookup"><span data-stu-id="b2233-138">?</span></span> |<span data-ttu-id="b2233-139">s?t localiza "sentou" e "sentir".</span><span class="sxs-lookup"><span data-stu-id="b2233-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="b2233-140">Qualquer sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="b2233-140">Any string of characters</span></span>| * |<span data-ttu-id="b2233-141">g\*s localiza gostar e gastar.</span><span class="sxs-lookup"><span data-stu-id="b2233-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="b2233-142">O início de uma palavra</span><span class="sxs-lookup"><span data-stu-id="b2233-142">The beginning of a word</span></span>|< |<span data-ttu-id="b2233-143"><(inter) localiza interseção e interessante, mas não localiza desinteresse.</span><span class="sxs-lookup"><span data-stu-id="b2233-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="b2233-144">O final de uma palavra</span><span class="sxs-lookup"><span data-stu-id="b2233-144">The end of a word</span></span> |> |<span data-ttu-id="b2233-145">(em)> localiza vargem e miragem, mas não localiza embrião.</span><span class="sxs-lookup"><span data-stu-id="b2233-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="b2233-146">Um dos caracteres especificados</span><span class="sxs-lookup"><span data-stu-id="b2233-146">One of the specified characters</span></span>|<span data-ttu-id="b2233-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="b2233-147">[ ]</span></span> |<span data-ttu-id="b2233-148">t[eo]m localiza tem e tom.</span><span class="sxs-lookup"><span data-stu-id="b2233-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="b2233-149">Qualquer caractere único deste intervalo</span><span class="sxs-lookup"><span data-stu-id="b2233-149">Any single character in this range</span></span>| <span data-ttu-id="b2233-150">[-]</span><span class="sxs-lookup"><span data-stu-id="b2233-150">[-]</span></span> |<span data-ttu-id="b2233-p109">[r-t]olo localiza rolo e solo. Os intervalos devem estar em ordem crescente.</span><span class="sxs-lookup"><span data-stu-id="b2233-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="b2233-153">Qualquer caractere único, exceto os caracteres do intervalo entre colchetes</span><span class="sxs-lookup"><span data-stu-id="b2233-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="b2233-155">t[!a-m]que localiza toque e trunque, mas não localiza taque ou tique.</span><span class="sxs-lookup"><span data-stu-id="b2233-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="b2233-156">Número de ocorrências exatas do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="b2233-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="b2233-157">{n}</span><span class="sxs-lookup"><span data-stu-id="b2233-157">{n}</span></span> |<span data-ttu-id="b2233-158">fe{2}d localiza feed, mas não fed.</span><span class="sxs-lookup"><span data-stu-id="b2233-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="b2233-159">Número mínimo de ocorrências do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="b2233-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="b2233-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="b2233-160">{n,}</span></span> |<span data-ttu-id="b2233-161">fe{1,}d localiza fed e feed.</span><span class="sxs-lookup"><span data-stu-id="b2233-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="b2233-162">De n a m ocorrências do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="b2233-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="b2233-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="b2233-163">{n,m}</span></span> |<span data-ttu-id="b2233-164">10{1,3} localiza 10, 100 e 1000.</span><span class="sxs-lookup"><span data-stu-id="b2233-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="b2233-165">Uma ou mais ocorrências do caractere ou expressão anterior</span><span class="sxs-lookup"><span data-stu-id="b2233-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="b2233-166">re@r localiza reter e reverter.</span><span class="sxs-lookup"><span data-stu-id="b2233-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="b2233-167">Escape de caracteres especiais</span><span class="sxs-lookup"><span data-stu-id="b2233-167">Escaping the special characters</span></span>

<span data-ttu-id="b2233-p110">A pesquisa com caracteres curinga é essencialmente igual à pesquisa em uma expressão regular. Há caracteres especiais em expressões regulares, como “[', ']”, “(', ')”, “{”, “}”, “\*”, “?”, “<”, “>”, “!” e “@”. Se um desses caracteres fizer parte da sequência de caracteres literal que o código está procurando, ele precisará usar o escape para que o Word saiba que ele deve ser tratado literalmente e não como parte da lógica da expressão regular. Para usar o escape em um caractere na pesquisa da interface de usuário do Word, prefixe-o com um caractere “\'”, mas, para usar o escape programaticamente, coloque-o entre caracteres “[]”. Por exemplo, “[\*]\*” pesquisa qualquer sequência de caracteres que comece com “\*” seguido por qualquer número de outros caracteres.</span><span class="sxs-lookup"><span data-stu-id="b2233-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="b2233-173">Exemplos</span><span class="sxs-lookup"><span data-stu-id="b2233-173">Examples</span></span>
<span data-ttu-id="b2233-174">Os exemplos a seguir demonstram cenários comuns.</span><span class="sxs-lookup"><span data-stu-id="b2233-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="b2233-175">Ignorar pesquisa de pontuação</span><span class="sxs-lookup"><span data-stu-id="b2233-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="b2233-176">Pesquisa com base em um prefixo</span><span class="sxs-lookup"><span data-stu-id="b2233-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="b2233-177">Pesquisa com base em um sufixo</span><span class="sxs-lookup"><span data-stu-id="b2233-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="b2233-178">Pesquisa usando caracteres curinga</span><span class="sxs-lookup"><span data-stu-id="b2233-178">Search using a wildcard</span></span>

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

<span data-ttu-id="b2233-179">Mais informações podem ser encontradas na [API de referência JavaScript do Word](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="b2233-179">More information can be found in the [Word JavaScript Reference API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).</span></span>