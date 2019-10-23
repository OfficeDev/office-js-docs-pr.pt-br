---
title: Obter o JavaScript IntelliSense no Visual Studio 2019
description: ''
ms.date: 10/11/2019
localization_priority: Priority
ms.openlocfilehash: afe2890015c853c2b70a00d94ea1b6a49296a14a
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626779"
---
# <a name="get-javascript-intellisense-in-visual-studio-2019"></a>Obter o JavaScript IntelliSense no Visual Studio 2019

Ao usar o Visual Studio 2019 para desenvolver Suplementos do Office, você pode usar o JSDoc para habilitar o IntelliSense para as variáveis, os objetos, os parâmetros e os valores de retorno de JavaScript. Este artigo fornece uma visão geral do JSDoc e como usá-lo para criar IntellSense no Visual Studio. Confira mais detalhes em [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) e [Suporte ao JSDoc no JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript). 

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
var subsetRange;
```
![IntelliSense para variável](../images/intellisense-vs17-var.png)

### <a name="parameter"></a>Parâmetro

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```
![IntelliSense para parâmetro](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>Valor de retorno

```js
/** @returns {Word.Range} */
function myFunc() {

}
```
![IntelliSense para valor de retorno](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>Tipos complexos

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```
![IntelliSense para tipo complexo](../images/intellisense-vs17-complex-type.png)

## <a name="see-also"></a>Confira também

- [Criar e depurar suplementos no Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md)
