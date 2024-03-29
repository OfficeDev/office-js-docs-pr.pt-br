---
title: Obter o JavaScript IntelliSense no Visual Studio
description: Saiba como usar o JSDoc para criar o IntelliSense para suas variáveis, objetos, parâmetros e valores retornados do JavaScript.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: deef6fe4356264534732e7f38a58a4079223686d
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889308"
---
# <a name="get-javascript-intellisense-in-visual-studio"></a>Obter o JavaScript IntelliSense no Visual Studio

Ao usar o Visual Studio 2019 e posterior para desenvolver Suplementos do Office, você pode usar o JSDoc para habilitar o IntelliSense para suas variáveis, objetos, parâmetros e valores retornados do JavaScript. Este artigo fornece uma visão geral do JSDoc e como usá-lo para criar IntellSense no Visual Studio. Confira mais detalhes em [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) e [Suporte ao JSDoc no JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).

## <a name="officejs-type-definitions"></a>Definições de tipo do Office.js

Você precisa fornecer as definições dos tipos no Office.js para o Visual Studio. Para fazer isso, é possível:

- Ter uma cópia local dos arquivos Office.js em uma pasta em sua solução denominada `\Office\1\`. Os modelos de projeto de Suplemento do Office no Visual Studio adicionam essa cópia local quando você cria o projeto de um suplemento.
- Use a versão online do Office.js adicionando um arquivo tsconfig.json à raiz do projeto de aplicativo da Web na solução do suplemento. O arquivo deve incluir o seguinte conteúdo:

    ```json
        {
            "compilerOptions": {
                "allowJs": true,            // These settings apply to JavaScript files also.
                "noEmit":  true             // Do not compile the JS (or TS) files in this project.
            },
            "exclude": [
                "node_modules",             // Don't include any JavaScript found under "node_modules".
                "Scripts/Office/1"          // Suppress loading all the JavaScript files from the Office NuGet package.
            ],
            "typeAcquisition": {
                "enable": true,             // Enable automatic fetching of type definitions for detected JavaScript libraries.
                "include": [ "office-js" ]  // Ensure that the "Office-js" type definition is fetched.
            }
        }
    ```

## <a name="jsdoc-syntax"></a>Sintaxe JSDoc

A técnica básica é incluir antes da variável (ou do parâmetro e assim por diante) um comentário que identifica seu tipo de dados. Isso permite que o IntelliSense no Visual Studio infira seus membros. Eis alguns exemplos:

### <a name="variable"></a>Variável

```js
/** @type {Excel.Range} */
let subsetRange;
```

![Trecho do IntelliSense para a variável 'subsetRange'.](../images/intellisense-vs22-var.png)

### <a name="parameter"></a>Parâmetro

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![Trecho do IntelliSense para o parâmetro 'paras' (parâmetro 'paragraphs' no exemplo de JavaScript).](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>Valor de retorno

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Trecho do IntelliSense para o valor retornado 'myFunc()'.](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>Tipos complexos

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![IntelliSense para declaração de tipo complexo de 'let myVar;', por exemplo.](../images/intellisense-vs22-complex-type.png)

## <a name="see-also"></a>Confira também

- [Desenvolver Suplementos do Office com o Visual Studio](develop-add-ins-visual-studio.md)
- [Depurar Suplementos do Office no Visual Studio](debug-office-add-ins-in-visual-studio.md)
