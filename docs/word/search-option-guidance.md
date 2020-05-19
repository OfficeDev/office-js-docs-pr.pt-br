---
title: Usar as opções de pesquisa para localizar o texto no suplemento do Word
description: Saiba como usar as opções de pesquisa no suplemento do Word
ms.date: 09/27/2019
localization_priority: Normal
ms.openlocfilehash: 1b0c1250b875ac2e61e68c65e9db6eba8fda4c67
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44276047"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>Usar as opções de pesquisa para localizar o texto no suplemento do Word

Os suplementos frequentemente precisam agir com base no texto de um documento.
Uma função de pesquisa é exposta por cada controle de conteúdo (isso inclui [Corpo](/javascript/api/word/word.body), [Parágrafo](/javascript/api/word/word.paragraph), [Intervalo](/javascript/api/word/word.range), [Tabela](/javascript/api/word/word.table), [ Coluna](/javascript/api/word/word.tablerow)e o objeto [ContentControl](/javascript/api/word/word.contentcontrol) base). Esta função assume uma cadeia de caracteres (ou expressão curinga) que representa o texto que você está procurando e um objeto [SearchOptions](/javascript/api/word/word.searchoptions). Retorna um conjunto de intervalos que correspondem ao texto de pesquisa.

## <a name="search-options"></a>Opções de pesquisa

As opções de pesquisa são uma coleção de valores boolianos que definem como o parâmetro de pesquisa deve ser tratado.

| Propriedade       | Descrição|
|:---------------|:----|
|ignorePunct|Obtém ou define um valor que indica se devem ser ignorados todos os caracteres de pontuação entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de pontuação" na caixa de diálogo Localizar e Substituir.|
|ignoreSpace|Obtém ou define um valor que indica se devem ser ignorados todos os espaços em branco entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de espaço em branco" na caixa de diálogo Localizar e Substituir.|
|matchCase|Obtém ou define um valor que determina quando realizar uma pesquisa que diferencia maiúsculas de minúsculas. Corresponde à caixa de seleção "Diferenciar maiúsculas de minúsculas", na caixa de diálogo "Localizar e Substituir" (menu Editar).|
|matchPrefix|Obtém ou define um valor que indica se se deve corresponder palavras que começam com a cadeia de caracteres de pesquisa. Corresponde à caixa de seleção "Corresponder prefixo" na caixa de diálogo "Localizar e Substituir".|
|matchSuffix|Obtém ou define um valor que indica se se deve corresponder palavras que terminam com a cadeia de caracteres de pesquisa. Corresponde à caixa de seleção "Corresponder sufixo", na caixa de diálogo "Localizar e Substituir".|
|matchWholeWord|Obtém ou define um valor que indica se a operação de localização localiza apenas palavras inteiras, e não texto que faz parte de uma palavra maior. Corresponde à caixa de seleção "Localizar apenas palavras inteiras" na caixa de diálogo Localizar e Substituir.|
|matchWildcards|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais. Corresponde à caixa de seleção "Usar caracteres curinga" na caixa de diálogo "Localizar e Substituir".|

## <a name="wildcard-guidance"></a>Diretrizes para caracteres curinga

A tabela a seguir fornece orientações em torno de caracteres curinga de pesquisa da API JavaScript do Word.

| Para localizar:         | Curinga |  Exemplo |
|:-----------------|:--------|:----------|
| Qualquer caractere simples| ? |c?l localiza "calor" e "caldo". |
|Qualquer cadeia de caracteres| * |g*s localiza gostar e gastar.|
|O início de uma palavra|< |< (inter) localiza interseção e interessante, mas não localiza desinteresse.|
|O final de uma palavra |> |(em)> localiza vargem e miragem, mas não localiza embrião.|
|Um dos caracteres especificados|[ ] |t[eo]m localiza tem e tom.|
|Qualquer caractere único deste intervalo| [-] |[r-t]olo localiza rolo e solo. Os intervalos devem estar em ordem crescente.|
|Qualquer caractere único, exceto os caracteres do intervalo entre colchetes|[!x-z] |t[!a-m]que localiza toque e trunque, mas não localiza taque ou tique.|
|Número de ocorrências exatas do caractere ou expressão anterior|{n} |ve{2}m localiza veem, mas não vem.|
|Número mínimo de ocorrências do caractere ou expressão anterior|{n,} |ve{1,}m localiza vem e veem.|
|Número de ocorrências do caractere ou expressão anterior dentro de um intervalo|{n,m} |10{1,3} localiza 10, 100 e 1000.|
|Uma ou mais ocorrências do caractere ou expressão anterior|@ |re@r localiza reter e reverter.|

### <a name="escaping-the-special-characters"></a>Escapar os caracteres especiais

A pesquisa com caracteres curinga é essencialmente igual à pesquisa em uma expressão regular. Há caracteres especiais em expressões regulares, como “[', ']”, “(', ')”, “{”, “}”, “\*”, “?”, “<”, “>”, “!” e “@”. Se um desses caracteres fizer parte da cadeia de caracteres literal que o código está procurando, ele precisará ser escapado para que o Word saiba que ele deve ser tratado literalmente e não como parte da lógica da expressão regular. Para escapar um caractere na pesquisa da interface de usuário do Word, prefixe-o com um caractere “\'”, mas, para escapá-lo programaticamente, coloque-o entre caracteres “[]”. Por exemplo, “[\*]\*” pesquisa qualquer cadeia de caracteres que comece com “\*” seguido por qualquer número de outros caracteres. 

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
    var searchResults = context.document.body.search('to*n', {matchWildcards: true});

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

Mais informações podem ser encontradas na [Referência de API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md).
