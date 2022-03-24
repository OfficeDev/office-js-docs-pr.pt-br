---
title: Teste de unidade em Office de complementos
description: Saiba como usar o código de teste de unidade que chama as OFFICE APIs JavaScript.
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6f6b0483b23c3f7199a8bd308bf8a4118402ee08
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746132"
---
# <a name="unit-testing-in-office-add-ins"></a>Teste de unidade em Office de complementos

Os testes de unidade verificam a funcionalidade do seu complemento sem exigir conexões de rede ou serviço, incluindo conexões com o Office aplicativo. O código do lado do servidor de teste de unidade e o código do  lado do cliente que não chama as [APIs JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md), é o mesmo em complementos do Office como está em qualquer aplicativo Web, portanto, não requer documentação especial. Mas o código do lado do cliente que chama as OFFICE APIs JavaScript é um desafio para testar. Para resolver esses problemas, criamos uma biblioteca para simplificar a criação de objetos Office simulados em testes de unidade: [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). A biblioteca facilita os testes das seguintes maneiras:

- As APIs do Office JavaScript devem ser inicializadas em um controle webview no contexto de um aplicativo Office (Excel, Word, etc.), portanto, não podem ser carregadas no processo em que os testes de unidade são executados no computador de desenvolvimento. A biblioteca Office-Addin-Mock pode ser importada para seus arquivos de teste, o que permite Office simulação de APIs JavaScript Office dentro do processo node.js no qual os testes são executados.
- As [APIs específicas](../develop/understanding-the-javascript-api-for-office.md#api-models) do aplicativo têm [métodos](../develop/application-specific-api-model.md#load) de carga e sincronização que devem ser chamados em uma ordem específica em relação a outras funções e umas às outras.[](../develop/application-specific-api-model.md#sync) Além disso, o `load` método deve ser chamado com determinados parâmetros, dependendo de quais propriedades Office objetos serão lidos por código posteriormente na função que está  sendo testada. Mas as estruturas de teste de unidade são inerentemente sem estado, `load` `sync` portanto, não podem manter um registro de se ou foi chamado ou para quais parâmetros foram passados para `load`. Os objetos simulados que você cria com a biblioteca Office-Addin-Mock têm estado interno que mantém o controle dessas coisas. Isso permite que os objetos mock emular o comportamento de erro de objetos Office reais. Por exemplo, se a `load`função que está sendo testada tentar ler uma propriedade que não foi passada pela primeira vez para , o teste retornará um erro semelhante ao que Office retornaria.

A biblioteca não depende das OFFICE JavaScript e pode ser usada com qualquer estrutura de teste de unidade JavaScript, como:

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmim](https://jasmine.github.io/)

Os exemplos neste artigo usam a estrutura Jest. Há exemplos usando a estrutura Mocha na [home page Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Pré-requisitos

Este artigo supõe que você está familiarizado com os conceitos básicos de teste de unidade e simulação, incluindo como criar e executar arquivos de teste e que você tem alguma experiência com uma estrutura de teste de unidade.

> [!TIP]
> Se você estiver trabalhando com Visual Studio, recomendamos que você leia o artigo Unit [testing JavaScript and TypeScript in Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) for some basic information about JavaScript unit testing in Visual Studio and then return to this article.

## <a name="install-the-tool"></a>Instalar a ferramenta

Para instalar a biblioteca, abra um prompt de comando, navegue até a raiz do projeto do seu complemento e insira o seguinte comando.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Uso básico

1. Seu projeto terá um ou mais arquivos de teste. (Consulte as instruções para sua estrutura de teste e os arquivos de teste de exemplo em Exemplos(#examples) abaixo.) Import the library, with the `require` or `import` keyword, to any test file that has a test of a function that calls the Office JavaScript APIs, as shown in the following example.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Importe o módulo que contém a função de complemento que você deseja testar com a palavra-chave `require` ou `import` . A seguir está um exemplo que supõe que seu arquivo de teste está em uma subpasta da pasta com os arquivos de código do seu complemento.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Crie um objeto de dados que tenha as propriedades e subpropropriedades que você precisa simular para testar a função. A seguir, um exemplo de um objeto que simula a propriedade [Excel Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) e o [método Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)). Este não é o objeto de simulação final. Pense nele como um objeto de semente que é usado para `OfficeMockObject` criar o objeto mock final.

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. Passe o objeto de dados para o `OfficeMockObject` construtor. Observe o seguinte sobre o objeto `OfficeMockObject` retornado.

   - É uma simulação simplificada de um [objeto OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) .
   - O objeto mock tem todos os membros do objeto de dados e também tem implementações simuladas dos `load` métodos e `sync` .
   - O objeto mock imitará o comportamento de erro crucial do `ClientRequestContext` objeto. Por exemplo, se a API de Office que você está testando tentar ler uma propriedade sem primeiro carregar a `sync`propriedade e chamar , o teste falhará com um erro semelhante ao que seria lançado no tempo de execução de produção: "Erro, propriedade não carregada".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > A documentação de referência completa do `OfficeMockObject` tipo está [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. Na sintaxe da estrutura de teste, adicione um teste da função. Use o `OfficeMockObject` objeto no lugar do objeto que ele simula, nesse caso, o `ClientRequestContext` objeto. O exemplo a seguir continua em Jest. Este teste `getSelectedRangeAddress`de exemplo pressupõe que a função de complemento que está sendo testada seja chamada , `ClientRequestContext` que ele tem um objeto como um parâmetro e que se destina a retornar o endereço do intervalo selecionado no momento. O exemplo completo é [posterior neste artigo](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Execute o teste de acordo com a documentação da estrutura de teste e suas ferramentas de desenvolvimento. Normalmente, há um arquivo **package.json** com um script que executa a estrutura de teste. Por exemplo, se Jest for a estrutura, **package.json** conteria o seguinte:

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   Para executar o teste, insira o seguinte em um prompt de comando na raiz do projeto.

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>Exemplos

Os exemplos nesta seção usam Jest com suas configurações padrão. Essas configurações suportam módulos CommonJS. Consulte a [documentação Jest sobre](https://jestjs.io/docs/getting-started) como configurar o Jest e o node.js para dar suporte a módulos ECMAScript e para dar suporte a TypeScript. Para executar qualquer um desses exemplos, execute as etapas a seguir.

1. Crie um Office de Office para o aplicativo host Office apropriado (por exemplo, Excel ou Word). Uma maneira de fazer isso rapidamente é usar o gerador [Yeoman para Office Add-ins](../develop/yeoman-generator-overview.md).
1. Na raiz do projeto, [instale Jest](https://jestjs.io/docs/getting-started).
1. [Instale a ferramenta office-addin-mock](#install-the-tool).
1. Crie um arquivo exatamente como o primeiro arquivo no exemplo e adicione-o à pasta que contém os outros arquivos de origem do projeto, geralmente chamados `\src`de .
1. Crie uma subpasta para a pasta de arquivo de origem e dê a ela um nome apropriado, como `\tests`.
1. Crie um arquivo exatamente como o arquivo de teste no exemplo e adicione-o à subpasta.
1. Adicione um `test` script ao **arquivo package.json** e execute o teste, conforme descrito em [Uso básico](#basic-usage).

### <a name="mocking-the-office-common-apis"></a>Simulando as OFFICE COMUNS

Este exemplo pressupara um Office para qualquer host que suporte as [APIs](../develop/office-javascript-api-object-model.md) comuns do Office (por exemplo, Excel, PowerPoint ou Word). O complemento tem um de seus recursos em um arquivo chamado `my-common-api-add-in-feature.js`. O seguinte mostra o conteúdo do arquivo. A `addHelloWorldText` função define o texto "Hello World!" para o que está selecionado no documento no momento; por exemplo; um intervalo no Word, ou uma célula em Excel, ou uma caixa de texto no PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

O arquivo de teste, nomeado `my-common-api-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do complemento. O seguinte mostra o conteúdo do arquivo. Observe que a propriedade de nível superior é `context`, um [Office. Objeto Context](/javascript/api/office/office.context), portanto, o objeto que está sendo simulado é o pai dessa propriedade: um [objeto Office](/javascript/api/office). Observe o seguinte sobre este código:

- O `OfficeMockObject` construtor não adiciona  todas as classes Office enum ao objeto mock`Office`, portanto, `CoercionType.Text` o valor referenciado no método de complemento deve ser adicionado explicitamente no objeto de semente.
- Como a Office JavaScript não é carregada no processo de nó, `Office` o objeto referenciado no código do complemento deve ser declarado e inicializado.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### <a name="mocking-the-outlook-apis"></a>Mocking the Outlook APIs

Embora estritamente falando, as APIs Outlook fazem parte do modelo de API Comum, elas têm uma arquitetura especial criada em torno do objeto [Mailbox](/javascript/api/outlook/office.mailbox), portanto, fornecemos um exemplo distinto para Outlook. Este exemplo assume um Outlook que tem um de seus recursos em um arquivo chamado `my-outlook-add-in-feature.js`. O seguinte mostra o conteúdo do arquivo. A `addHelloWorldText` função define o texto "Hello World!" para o que está selecionado no momento na janela de composição de mensagem.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

O arquivo de teste, nomeado `my-outlook-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do complemento. O seguinte mostra o conteúdo do arquivo. Observe que a propriedade de nível superior é `context`, um [Office. Objeto Context](/javascript/api/office/office.context), portanto, o objeto que está sendo simulado é o pai dessa propriedade: um [objeto Office](/javascript/api/office). Observe o seguinte sobre este código:

- A `host` propriedade no objeto mock é usada internamente pela biblioteca simulada para identificar o Office aplicativo. É obrigatório para Outlook. Atualmente, ele não serve para nenhuma outra finalidade Office aplicativo.
- Como a Office JavaScript não é carregada no processo de nó, `Office` o objeto referenciado no código do complemento deve ser declarado e inicializado.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### <a name="mocking-the-office-application-specific-apis"></a>Simulando as APIs Office específicas do aplicativo

Quando você estiver testando funções que usam AS APIs específicas do aplicativo, certifique-se de que você está simulando o tipo certo de objeto. Há duas opções:

- Mock a [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Faça isso quando a função que está sendo testada atende a ambas as seguintes condições:

  - Ele não chama um *Host*.`run` como [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Ele não faz referência a nenhuma outra propriedade direta ou método de um *objeto Host* .

- Mock a *Host* object, such as [Excel](/javascript/api/excel) or [Word](/javascript/api/word). Faça isso quando a opção anterior não for possível.

Exemplos de ambos os tipos de testes estão nas subseções abaixo.

#### <a name="mocking-a-clientrequestcontext-object"></a>Simulando um objeto ClientRequestContext

Este exemplo assume um Excel que tem um de seus recursos em um arquivo chamado `my-excel-add-in-feature.js`. O seguinte mostra o conteúdo do arquivo. Observe que o `getSelectedRangeAddress` é um método auxiliar chamado dentro do retorno de chamada que é passado para `Excel.run`.

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

O arquivo de teste, nomeado `my-excel-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do complemento. O seguinte mostra o conteúdo do arquivo. Observe que a propriedade de nível superior é `workbook`, portanto, o objeto que está sendo simulado é o pai de um `Excel.Workbook`: um `ClientRequestContext` objeto.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### <a name="mocking-a-host-object"></a>Simulando um objeto host

Este exemplo assume um complemento do Word que tem um de seus recursos em um arquivo chamado `my-word-add-in-feature.js`. O seguinte mostra o conteúdo do arquivo.

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

O arquivo de teste, nomeado `my-word-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do complemento. O seguinte mostra o conteúdo do arquivo. Observe que a `ClientRequestContext` propriedade de nível superior `context`é , um objeto, portanto, o objeto que está sendo simulado é o pai dessa propriedade: um `Word` objeto. Observe o seguinte sobre este código:

- Quando o `OfficeMockObject` construtor criar o objeto mock final, ele garantirá que o objeto filho `ClientRequestContext` tenha `sync` e `load` métodos.
- O `OfficeMockObject` construtor não *adiciona um* método ao `run` objeto mock `Word` , portanto, ele deve ser adicionado explicitamente no objeto de semente.
- O `OfficeMockObject` construtor não adiciona  todas as classes de número do Word ao objeto mock`Word`, portanto, `InsertLocation.end` o valor referenciado no método de complemento deve ser adicionado explicitamente no objeto de semente.
- Como a Office JavaScript não é carregada no processo de nó, `Word` o objeto referenciado no código do complemento deve ser declarado e inicializado.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run method.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

> [!NOTE]
> A documentação de referência completa do `OfficeMockObject` tipo está [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>Confira também

- [Office-Addin-Mock ponto de instalação da página npm](https://www.npmjs.com/package/office-addin-mock). 
- O repo de código aberto [é Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmim](https://jasmine.github.io/)
