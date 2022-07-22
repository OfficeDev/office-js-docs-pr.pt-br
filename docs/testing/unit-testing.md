---
title: Teste de unidade em Suplementos do Office
description: Saiba como fazer o teste de unidade de código que chama as APIs JavaScript do Office.
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21858a68734ca5d07621f3e9c88b147ebac7dde6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958745"
---
# <a name="unit-testing-in-office-add-ins"></a>Teste de unidade em Suplementos do Office

Os testes de unidade verificam a funcionalidade do suplemento sem exigir conexões de rede ou serviço, incluindo conexões com o aplicativo do Office. O código do lado do servidor de teste de unidade e o código do  lado do cliente que não chama as [APIs JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md) são os mesmos em Suplementos do Office que em qualquer aplicativo Web, portanto, não requer documentação especial. Mas o código do lado do cliente que chama as APIs JavaScript do Office é um desafio para testar. Para resolver esses problemas, criamos uma biblioteca para simplificar a criação de objetos fictícios do Office em testes de unidade: [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). A biblioteca facilita o teste das seguintes maneiras:

- As APIs JavaScript do Office devem ser inicializadas em um controle de modo de exibição da Web no contexto de um aplicativo do Office (Excel, Word etc.), para que não possam ser carregadas no processo em que os testes de unidade são executados no computador de desenvolvimento. A biblioteca Office-Addin-Mock pode ser importada para seus arquivos de teste, o que permite a simulação de APIs JavaScript do Office dentro do processo node.js no qual os testes são executados.
- As [APIs específicas do](../develop/understanding-the-javascript-api-for-office.md#api-models) aplicativo têm [métodos](../develop/application-specific-api-model.md#load) de carregamento e sincronização que devem ser chamados em uma ordem específica em relação a outras funções e entre si.[](../develop/application-specific-api-model.md#sync) Além disso, o `load` método deve ser chamado com determinados parâmetros, dependendo de quais propriedades de objetos do Office serão lidas pelo código posteriormente na função que está sendo testada. Mas as estruturas de teste de unidade são inerentemente sem estado, portanto, `load` `sync` elas não podem manter um registro de se ou foi chamado ou para quais parâmetros foram passados `load`. Os objetos fictícios criados com a biblioteca Office-Addin-Mock têm um estado interno que controla essas coisas. Isso permite que os objetos fictícios emularem o comportamento de erro de objetos reais do Office. Por exemplo, se a `load`função que está sendo testada tentar ler uma propriedade que não foi passada pela primeira vez, o teste retornará um erro semelhante ao que o Office retornaria.

A biblioteca não depende das APIs JavaScript do Office e pode ser usada com qualquer estrutura de teste de unidade JavaScript, como:

- [Brincadeira](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmine](https://jasmine.github.io/)

Os exemplos neste artigo usam a estrutura Jest. Há exemplos usando a estrutura Mocha na [home page do Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Pré-requisitos

Este artigo pressupõe que você esteja familiarizado com os conceitos básicos de teste de unidade e simulação, incluindo como criar e executar arquivos de teste e que você tem alguma experiência com uma estrutura de teste de unidade.

> [!TIP]
> Se você estiver trabalhando com o Visual Studio, recomendamos que leia o artigo Teste de unidade [JavaScript e TypeScript no Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) para obter algumas informações básicas sobre o teste de unidade javaScript no Visual Studio e, em seguida, retornar a este artigo.

## <a name="install-the-tool"></a>Instalar a ferramenta

Para instalar a biblioteca, abra um prompt de comando, navegue até a raiz do projeto de suplemento e insira o comando a seguir.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Uso básico

1. Seu projeto terá um ou mais arquivos de teste. (Consulte as instruções para a estrutura de teste e os arquivos de teste de exemplo em Exemplos (#examples) abaixo.) Importe a biblioteca, `require` `import` com a palavra-chave ou a palavra-chave, para qualquer arquivo de teste que tenha um teste de uma função que chame as APIs JavaScript do Office, conforme mostrado no exemplo a seguir.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Importe o módulo que contém a função de suplemento que você deseja testar com a palavra-chave `require` ou a palavra-chave `import` . A seguir está um exemplo que pressupõe que o arquivo de teste esteja em uma subpasta da pasta com os arquivos de código do suplemento.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Crie um objeto de dados que tenha as propriedades e subpropriedades que você precisa simular para testar a função. A seguir está um exemplo de um objeto que simula a propriedade [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) do Excel e o [método Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) . Este não é o objeto de simulação final. Pense nele como um objeto de semente usado para `OfficeMockObject` criar o objeto fictício final.

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
   - O objeto fictício tem todos os membros do objeto de dados e também tem implementações fict a e `load` `sync` métodos.
   - O objeto fictício imitará o comportamento de erro crucial do `ClientRequestContext` objeto. Por exemplo, se a API do Office que você está testando tentar ler uma propriedade sem primeiro carregar a `sync`propriedade e chamar, o teste falhará com um erro semelhante ao que seria gerado no runtime de produção: "Erro, propriedade não carregada".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > A documentação de referência completa `OfficeMockObject` para o tipo está [no Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. Na sintaxe da estrutura de teste, adicione um teste da função. Use o `OfficeMockObject` objeto no lugar do objeto que ele simula, nesse caso, o `ClientRequestContext` objeto. O exemplo a seguir continua em Jest. Este teste `getSelectedRangeAddress`de exemplo pressupõe que a função de suplemento que está sendo testada é chamada, `ClientRequestContext` que ele usa um objeto como um parâmetro e que se destina a retornar o endereço do intervalo selecionado no momento. O exemplo completo [é posteriormente neste artigo](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Execute o teste de acordo com a documentação da estrutura de teste e suas ferramentas de desenvolvimento. Normalmente, há um arquivo **package.json** com um script que executa a estrutura de teste. Por exemplo, se Jest for a estrutura, **package.json** conterá o seguinte:

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

Os exemplos nesta seção usam Jest com suas configurações padrão. Essas configurações dão suporte a módulos CommonJS. Consulte a [documentação do Jest](https://jestjs.io/docs/getting-started) para saber como configurar o Jest e o node.js para dar suporte a módulos ECMAScript e para dar suporte ao TypeScript. Para executar qualquer um desses exemplos, execute as etapas a seguir.

1. Crie um projeto de suplemento do Office para o aplicativo host do Office apropriado (por exemplo, Excel ou Word). Uma maneira de fazer isso rapidamente é usar o [gerador Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).
1. Na raiz do projeto, [instale o Jest](https://jestjs.io/docs/getting-started).
1. [Instale a ferramenta office-addin-mock](#install-the-tool).
1. Crie um arquivo exatamente como o primeiro arquivo no exemplo e adicione-o à pasta que contém os outros arquivos de origem do projeto, geralmente chamados de `\src`.
1. Crie uma subpasta para a pasta do arquivo de origem e dê a ela um nome apropriado, como `\tests`.
1. Crie um arquivo exatamente como o arquivo de teste no exemplo e adicione-o à subpasta.
1. Adicione um `test` script ao **arquivo package.json** e execute o teste, conforme descrito [em Uso básico](#basic-usage).

### <a name="mocking-the-office-common-apis"></a>Simulando as APIs comuns do Office

Este exemplo pressupõe um Suplemento do Office para qualquer host que dê suporte às [APIs](../develop/office-javascript-api-object-model.md) comuns do Office (por exemplo, Excel, PowerPoint ou Word). O suplemento tem um de seus recursos em um arquivo chamado `my-common-api-add-in-feature.js`. O exemplo a seguir mostra o conteúdo do arquivo. A `addHelloWorldText` função define o texto "Olá, Mundo!" para o que está selecionado no momento no documento; por exemplo; um intervalo no Word ou uma célula no Excel ou uma caixa de texto no PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

O arquivo de teste, nomeado `my-common-api-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do suplemento. O exemplo a seguir mostra o conteúdo do arquivo. Observe que a propriedade de `context`nível superior é , um objeto [Office.Context](/javascript/api/office/office.context) , portanto, o objeto que está sendo simulado é o pai dessa propriedade: um [objeto do Office](/javascript/api/office) . Observe o seguinte sobre este código:

- O `OfficeMockObject` construtor não adiciona  todas as classes de enumeração `Office` do Office ao objeto fictício, portanto, `CoercionType.Text` o valor referenciado no método de suplemento deve ser adicionado explicitamente no objeto de semente.
- Como a biblioteca JavaScript do Office não está carregada no processo de nó, `Office` o objeto referenciado no código do suplemento deve ser declarado e inicializado.

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

### <a name="mocking-the-outlook-apis"></a>Simulando as APIs do Outlook

Embora estritamente falando, as APIs do Outlook fazem parte do modelo de API Comum, elas têm uma arquitetura especial criada em torno do objeto [Mailbox](/javascript/api/outlook/office.mailbox) , portanto, fornecemos um exemplo distinto para o Outlook. Este exemplo pressupõe um Outlook que tenha um de seus recursos em um arquivo chamado `my-outlook-add-in-feature.js`. O exemplo a seguir mostra o conteúdo do arquivo. A `addHelloWorldText` função define o texto "Olá, Mundo!" para o que estiver selecionado no momento na janela de composição da mensagem.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

O arquivo de teste, nomeado `my-outlook-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do suplemento. O exemplo a seguir mostra o conteúdo do arquivo. Observe que a propriedade de `context`nível superior é , um objeto [Office.Context](/javascript/api/office/office.context) , portanto, o objeto que está sendo simulado é o pai dessa propriedade: um [objeto do Office](/javascript/api/office) . Observe o seguinte sobre este código:

- A `host` propriedade no objeto fictício é usada internamente pela biblioteca simulada para identificar o aplicativo do Office. É obrigatório para o Outlook. Atualmente, ele não serve para nenhum outro aplicativo do Office.
- Como a biblioteca JavaScript do Office não está carregada no processo de nó, `Office` o objeto referenciado no código do suplemento deve ser declarado e inicializado.

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

### <a name="mocking-the-office-application-specific-apis"></a>Simulando as APIs específicas do aplicativo do Office

Ao testar funções que usam as APIs específicas do aplicativo, certifique-se de que você está simulando o tipo correto de objeto. Há duas opções:

- Simular [um OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Faça isso quando a função que está sendo testada atender às duas condições a seguir:

  - Ele não chama um *host*.`run` função, como [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Ele não faz referência a nenhuma outra propriedade direta ou método de um *objeto Host* .

- Simular *um objeto* Host, como [o Excel](/javascript/api/excel) ou [o Word](/javascript/api/word). Faça isso quando a opção anterior não for possível.

Exemplos de ambos os tipos de testes estão nas subseções abaixo.

#### <a name="mocking-a-clientrequestcontext-object"></a>Simulando um objeto ClientRequestContext

Este exemplo pressupõe um suplemento do Excel que tenha um de seus recursos em um arquivo chamado `my-excel-add-in-feature.js`. O exemplo a seguir mostra o conteúdo do arquivo. Observe que é `getSelectedRangeAddress` um método auxiliar chamado dentro do retorno de chamada que é passado para `Excel.run`.

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

O arquivo de teste, nomeado `my-excel-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do suplemento. O exemplo a seguir mostra o conteúdo do arquivo. Observe que a propriedade de nível superior é `workbook`, portanto, o objeto que está sendo simulado é o pai de um `Excel.Workbook`: um `ClientRequestContext` objeto.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
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

Este exemplo pressupõe um suplemento do Word que tenha um de seus recursos em um arquivo chamado `my-word-add-in-feature.js`. O exemplo a seguir mostra o conteúdo do arquivo.

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

O arquivo de teste, nomeado `my-word-add-in-feature.test.js` está em uma subpasta, em relação ao local do arquivo de código do suplemento. O exemplo a seguir mostra o conteúdo do arquivo. Observe que a propriedade de nível superior `context`é , `ClientRequestContext` um objeto, portanto, o objeto que está sendo simulado é o pai dessa propriedade: um `Word` objeto. Observe o seguinte sobre este código:

- Quando o `OfficeMockObject` construtor cria o objeto de simulação final, ele garantirá que o objeto filho `ClientRequestContext` tenha `sync` e `load` métodos.
- O `OfficeMockObject` construtor não *adiciona uma* função ao `run` objeto fictício `Word` , portanto, ele deve ser adicionado explicitamente no objeto de semente.
- O `OfficeMockObject` construtor não adiciona  todas as classes de enumeração `Word` do Word ao objeto fictício, portanto, `InsertLocation.end` o valor referenciado no método de suplemento deve ser adicionado explicitamente no objeto de semente.
- Como a biblioteca JavaScript do Office não está carregada no processo de nó, `Word` o objeto referenciado no código do suplemento deve ser declarado e inicializado.

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
  // Mock the Word.run function.
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
> A documentação de referência completa `OfficeMockObject` para o tipo está [no Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>Confira também

- [Ponto de instalação da página npm do Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock) . 
- O código aberto repositório é [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Brincadeira](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Jasmine](https://jasmine.github.io/)
