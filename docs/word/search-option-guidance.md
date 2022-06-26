---
title: Usar opções de pesquisa no suplemento do Word para localizar texto
description: Saiba como usar as opções de pesquisa em seu suplemento do Word.
ms.date: 02/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 988349275dc350a342dfcb80e8e999c76de78e7d
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229705"
---
# <a name="use-search-options-in-your-word-add-in-to-find-text"></a>Usar opções de pesquisa no suplemento do Word para localizar texto

Os suplementos frequentemente precisam agir com base no texto de um documento. Uma função de pesquisa é exposta por cada controle de conteúdo (isso inclui [Corpo](/javascript/api/word/word.body), [Parágrafo](/javascript/api/word/word.paragraph), [Intervalo](/javascript/api/word/word.range), [Tabela](/javascript/api/word/word.table), [ Coluna](/javascript/api/word/word.tablerow)e o objeto [ContentControl](/javascript/api/word/word.contentcontrol) base). Esta função assume uma cadeia de caracteres (ou expressão curinga) que representa o texto que você está procurando e um objeto [SearchOptions](/javascript/api/word/word.searchoptions). Retorna um conjunto de intervalos que correspondem ao texto de pesquisa.

## <a name="search-options"></a>Opções de pesquisa

As opções de pesquisa são uma coleção de valores boolianos que definem como o parâmetro de pesquisa deve ser tratado.

| Propriedade       | Descrição|
|:---------------|:----|
|ignorePunct|Obtém ou define um valor que indica se devem ser ignorados todos os caracteres de pontuação entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de pontuação" na caixa de **diálogo Localizar e** Substituir.|
|ignoreSpace|Obtém ou define um valor que indica se devem ser ignorados todos os espaços em branco entre as palavras. Corresponde à caixa de seleção "Ignorar caracteres de espaço em branco" na caixa de **diálogo Localizar e** Substituir.|
|matchCase|Obtém ou define um valor que indica se uma pesquisa que diferencia maiúsculas de minúsculas deve ser executada. Corresponde à caixa de seleção "Diferenciar maiúsculas e minúsculas" na caixa de **diálogo Localizar e** Substituir.|
|matchPrefix|Obtém ou define um valor que indica se se deve corresponder palavras que começam com a cadeia de caracteres de pesquisa. Corresponde à caixa de seleção "Coincidir prefixo" na caixa de **diálogo Localizar e** Substituir.|
|matchSuffix|Obtém ou define um valor que indica se se deve corresponder palavras que terminam com a cadeia de caracteres de pesquisa. Corresponde à caixa de seleção "Coincidir sufixo" na caixa de **diálogo Localizar e** Substituir.|
|matchWholeWord|Obtém ou define um valor que indica se a operação de localização localiza apenas palavras inteiras, e não texto que faz parte de uma palavra maior. Corresponde à caixa de seleção "Localizar somente palavras inteiras" na caixa de **diálogo Localizar e** Substituir.|
|matchWildcards|Obtém ou define um valor que indica se a pesquisa será realizada com operadores de pesquisa especiais. Corresponde à caixa de seleção "Usar caracteres curinga" na caixa de **diálogo Localizar e** Substituir.|

## <a name="wildcard-guidance"></a>Diretrizes para caracteres curinga

A tabela a seguir fornece orientações em torno de caracteres curinga de pesquisa da API JavaScript do Word.

| Para localizar:         | Curinga |  Exemplo |
|:-----------------|:--------|:----------|
|Qualquer caractere simples| ? |c?l localiza "calor" e "caldo". |
|Qualquer cadeia de caracteres| * |g*s localiza gostar e gastar.|
|O início de uma palavra|< |< (inter) localiza interseção e interessante, mas não localiza desinteresse.|
|O final de uma palavra |> |(em)> localiza vargem e miragem, mas não localiza embrião.|
|Um dos caracteres especificados|[ ] |t[eo]m localiza tem e tom.|
|Qualquer caractere único deste intervalo| [-] |[r-t]olo localiza rolo e solo. Os intervalos devem estar em ordem crescente.|
|Qualquer caractere único, exceto os caracteres do intervalo entre colchetes|[!x-z] |t[!a-m]que localiza toque e trunque, mas não localiza taque ou tique.|
|Exatamente *n* ocorrências do caractere ou expressão anterior|{n} |ve{2}m localiza veem, mas não vem.|
|Pelo menos *n* ocorrências do caractere ou expressão anterior|{n,} |ve{1,}m localiza vem e veem.|
|De *n* a *m* ocorrências do caractere ou expressão anterior|{n,m} |10{1,3} localiza 10, 100 e 1000.|
|Uma ou mais ocorrências do caractere ou expressão anterior|@ |re@r localiza reter e reverter.|

### <a name="escaping-special-characters"></a>Escape de caracteres especiais

A pesquisa com caracteres curinga é essencialmente igual à pesquisa em uma expressão regular. Há caracteres especiais em expressões regulares, como “[', ']”, “(', ')”, “{”, “}”, “\*”, “?”, “<”, “>”, “!” e “@”. Se um desses caracteres fizer parte da cadeia de caracteres literal que o código está procurando, ele precisará ser escapado para que o Word saiba que ele deve ser tratado literalmente e não como parte da lógica da expressão regular. Para escapar de um caractere na pesquisa de interface do usuário do Word, você o precederia com um caractere de barra invertida ('\\'), mas, para escapar programaticamente, coloque-o entre caracteres '[]'. Por exemplo, “[\*]\*” pesquisa qualquer cadeia de caracteres que comece com “\*” seguido por qualquer número de outros caracteres.

## <a name="examples"></a>Exemplos

Os exemplos a seguir demonstram cenários comuns.

### <a name="ignore-punctuation-search"></a>Ignorar pesquisa de pontuação

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-prefix"></a>Pesquisa com base em um prefixo

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-suffix"></a>Pesquisa com base em um sufixo

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-using-a-wildcard"></a>Pesquisa usando caracteres curinga

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

Mais informações podem ser encontradas na [Referência de API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md).
