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
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Usar as opções de pesquisa para encontrar texto no seu suplemento do Word 

Os suplementos frequentemente precisam agir com base no texto de um documento. Uma função de pesquisa é exposta por todo controle de conteúdo (Isso inclui [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js), [Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js), [Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js), [Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js), [TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js) e o objeto base [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js)). Essa função usa uma sequência de caracteres (ou expressão curinga) que representa o texto que você está procurando e um objeto [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js). Ela retorna uma coleção de intervalos que correspondem ao texto da pesquisa.

## <a name="search-options"></a>Opções de pesquisa
As opções de pesquisa são uma coleção de valores booleanos que definem como o parâmetro de pesquisa deve ser tratado. 

| Propriedade     | Descrição|
|:---------------|:----|
|ignorePunct|Obtém ou define um valor indicando se deve ignorar todos os caracteres de pontuação entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de pontuação", na caixa de diálogo "Localizar e substituir".|
|ignoreSpace|Obtém ou define um valor indicando se deve ignorar todos os espaços em branco entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de espaço em branco", na caixa de diálogo "Localizar e substituir".|
|matchCase|Obtém ou define um valor indicando se deve ser realizada uma pesquisa que diferencia maiúsculas de minúsculas. Corresponde à caixa de seleção "Diferenciar maiúsculas/minúsculas" na caixa de diálogo "Localizar e substituir".|
|matchPrefix|Obtém ou define um valor que determina quando fazer correspondência com as palavras que começam com a sequência de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir prefixo", na caixa de diálogo "Localizar e substituir".|
|matchSuffix|Obtém ou define um valor que determina quando fazer correspondência com as palavras que terminam com a sequência de caracteres da pesquisa. Corresponde à caixa de seleção "Coincidir sufixo", na caixa de diálogo "Localizar e substituir".|
|matchWholeWord|Obtém ou define um valor indicando a operação para localizar somente palavras inteiras, não o texto que faz parte de uma palavra maior. Corresponde à caixa de seleção "Localizar apenas palavras inteiras", na caixa de diálogo "Localizar e substituir".|
|matchWildcards|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais. Corresponde à caixa de seleção "Usar caracteres curinga" na caixa de diálogo "Localizar e substituir".|

## <a name="wildcard-guidance"></a>Diretrizes para caracteres curinga
A tabela a seguir fornece diretrizes sobre os curingas de pesquisa da API JavaScript do Word.

| Para localizar:         | Curinga |  Exemplo |
|:-----------------|:--------|:----------|
| Qualquer caractere simples| ? |s?t localiza "sentou" e "sentir". |
|Qualquer sequência de caracteres| * |g*s localiza gostar e gastar.|
|O início de uma palavra|< |<(inter) localiza interseção e interessante, mas não localiza desinteresse.|
|O final de uma palavra |> |(em)> localiza vargem e miragem, mas não localiza embrião.|
|Um dos caracteres especificados|[ ] |t[eo]m localiza tem e tom.|
|Qualquer caractere único deste intervalo| [-] |[r-t]olo localiza rolo e solo. Os intervalos devem estar em ordem crescente.|
|Qualquer caractere único, exceto os caracteres do intervalo entre colchetes|[!x-z] |t[!a-m]que localiza toque e trunque, mas não localiza taque ou tique.|
|Número de ocorrências exatas do caractere ou expressão anterior|{n} |fe{2}d localiza feed, mas não fed.|
|Número mínimo de ocorrências do caractere ou expressão anterior|{n,} |fe{1,}d localiza fed e feed.|
|De n a m ocorrências do caractere ou expressão anterior|{n,m} |10{1,3} localiza 10, 100 e 1000.|
|Uma ou mais ocorrências do caractere ou expressão anterior|@ |re@r localiza reter e reverter.|

### <a name="escaping-the-special-characters"></a>Escape de caracteres especiais

A pesquisa com caracteres curinga é essencialmente igual à pesquisa em uma expressão regular. Há caracteres especiais em expressões regulares, como “[', ']”, “(', ')”, “{”, “}”, “\*”, “?”, “<”, “>”, “!” e “@”. Se um desses caracteres fizer parte da sequência de caracteres literal que o código está procurando, ele precisará usar o escape para que o Word saiba que ele deve ser tratado literalmente e não como parte da lógica da expressão regular. Para usar o escape em um caractere na pesquisa da interface de usuário do Word, prefixe-o com um caractere “\'”, mas, para usar o escape programaticamente, coloque-o entre caracteres “[]”. Por exemplo, “[\*]\*” pesquisa qualquer sequência de caracteres que comece com “\*” seguido por qualquer número de outros caracteres. 

## <a name="examples"></a>Exemplos
Os exemplos a seguir demonstram cenários comuns.

### <a name="ignore-punctuation-search"></a>Ignorar pesquisa de pontuação

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

### <a name="search-based-on-a-prefix"></a>Pesquisa com base em um prefixo

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

### <a name="search-based-on-a-suffix"></a>Pesquisa com base em um sufixo

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

### <a name="search-using-a-wildcard"></a>Pesquisa usando caracteres curinga

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

Mais informações podem ser encontradas na [API de referência JavaScript do Word](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).