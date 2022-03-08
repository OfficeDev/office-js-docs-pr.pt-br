---
title: Chamar funções internas de planilha do Excel usando as APIs JavaScript do Excel
description: Saiba como chamar funções de Excel de planilhas `VLOOKUP` `SUM`, como e usando Excel API JavaScript.
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: b6e3ca14064934ed45228d14a95a0226a998937c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341034"
---
# <a name="call-built-in-excel-worksheet-functions"></a>Chamar funções internas de planilha do Excel

Este artigo explica como chamar funções internas de planilha do Excel, como `VLOOKUP` e `SUM`, usando as API JavaScript do Excel. Também fornece a lista completa de funções internas de planilha Excel que podem ser chamadas usando as APIs JavaScript do Excel.

> [!NOTE]
> Para saber mais sobre como criar *funções personalizadas* no Excel usando as APIs JavaScript do Excel, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Chamar uma função de planilha

O trecho de código a seguir mostra como chamar uma função de planilha, onde `sampleFunction()` é um espaço reservado que deve ser substituído pelo nome da função a chamar e os parâmetros de entrada que a função exige. A `value` propriedade do objeto `FunctionResult` retornado por uma função de planilha contém o resultado da função especificada. Como este exemplo mostra, você deve `load` ter a `value` propriedade do `FunctionResult` objeto antes de poder lê-lo. Neste exemplo, o resultado da função está simplesmente sendo gravado no console.

```js
await Excel.run(async (context) => {
    let functionResult = context.workbook.functions.sampleFunction();
    functionResult.load('value');

    await context.sync();
    console.log('Result of the function: ' + functionResult.value);
});
```

> [!TIP]
> Confira a seção [Funções de planilha com suporte](#supported-worksheet-functions) deste artigo para obter uma lista das funções que podem ser chamadas usando as APIs JavaScript do Excel.

## <a name="sample-data"></a>Dados de exemplo

A imagem a seguir mostra uma tabela em uma planilha do Excel com dados de vendas para vários tipos de ferramentas durante um período de três meses. Cada número da tabela representa o número de unidades vendidas de uma ferramenta específica em um mês específico. Os exemplos a seguir mostram como aplicar funções internas da planilha nesses dados.

![Captura de tela dos dados de vendas no Excel para Hammer, Chave Inglesa e Saw nos meses de novembro, dezembro e janeiro.](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Exemplo 1: função individual

O exemplo a seguir se aplica à função `VLOOKUP` para os dados de exemplo descritos anteriormente a fim de identificar o número de chaves inglesas vendidas em novembro.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
});
```

## <a name="example-2-nested-functions"></a>Exemplo 2: funções aninhadas

O exemplo de código a seguir aplica a função `VLOOKUP` nos dados de amostras descritos anteriormente para identificar o número de chaves inglesas vendidas em novembro e em dezembro e, em seguida, aplica a função `SUM` para calcular o total de chaves inglesas vendido durante esses dois meses.

Como mostra este exemplo, quando uma ou mais chamadas de função são aninhadas dentro de outra chamada de função, você só precisa dar `load` no resultado final caso você queira ler (neste exemplo, `sumOfTwoLookups`). Os resultados intermediários (neste exemplo, o resultado de cada função `VLOOKUP`) serão calculados e usados para calcular o resultado final.

```js
await Excel.run(async (context) => {
    let range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    let sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false),
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    await context.sync();
    console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
});
```

## <a name="supported-worksheet-functions"></a>Funções de planilha com suporte

As seguintes funções internas de planilhas do Excel podem ser chamadas usando as APIs JavaScript do Excel.

| Função | Descrição |
|:---------------|:-----------|
| [Função ABS](https://support.microsoft.com/office/3420200f-5628-4e8c-99da-c99d7c87713c) | Retorna o valor absoluto de um número |
| [Função JUROSACUM](https://support.microsoft.com/office/fe45d089-6722-4fb3-9379-e1f911d8dc74) | Retorna juros acumulados de um título que paga juros periódicos |
| [Função JUROSACUMV](https://support.microsoft.com/office/f62f01f9-5754-4cc4-805b-0e70199328a7) | Retorna juros acumulados de um título que paga juros no vencimento |
| [Função ACOS](https://support.microsoft.com/office/cb73173f-d089-4582-afa1-76e5524b5d5b) | Retorna o arco cosseno de um número |
| [Função ACOSH](https://support.microsoft.com/office/e3992cc1-103f-4e72-9f04-624b9ef5ebfe) | Retorna o cosseno hiperbólico inverso de um número |
| [Função ACOT](https://support.microsoft.com/office/dc7e5008-fe6b-402e-bdd6-2eea8383d905) | Retorna o arco cotangente de um número |
| [Função ACOTH](https://support.microsoft.com/office/cc49480f-f684-4171-9fc5-73e4e852300f) | Retorna o arco cotangente hiperbólico de um número |
| [Função AMORDEGRC](https://support.microsoft.com/office/a14d0ca1-64a4-42eb-9b3d-b0dededf9e51) | Retorna a depreciação para cada período contábil usando o coeficiente de depreciação |
| [Função AMORLINC](https://support.microsoft.com/office/7d417b45-f7f5-4dba-a0a5-3451a81079a8) | Retorna a depreciação para cada período contábil |
| [Função E](https://support.microsoft.com/office/5f19b2e8-e1df-4408-897a-ce285a19e9d9) | Retorna `TRUE` se todos os seus argumentos forem verdadeiros |
| [Função ARÁBICO](https://support.microsoft.com/office/9a8da418-c17b-4ef9-a657-9370a30a674f) | Converte um número romano em arábico, como um número |
| [Função ÁREAS](https://support.microsoft.com/office/8392ba32-7a41-43b3-96b0-3695d2ec6152) | Retorna o número de áreas em uma referência |
| [Função ASC](https://support.microsoft.com/office/0b6abf1c-c663-4004-a964-ebc00b723266) | Altera letras do inglês ou katakana de largura total (bytes duplos) dentro de uma cadeia de caracteres para caracteres de meia largura (byte único) |
| [Função ASEN](https://support.microsoft.com/office/81fb95e5-6d6f-48c4-bc45-58f955c6d347) | Retorna o arco seno de um número |
| [Função ASENH](https://support.microsoft.com/office/4e00475a-067a-43cf-926a-765b0249717c) | Retorna o seno hiperbólico inverso de um número |
| [Função ATAN](https://support.microsoft.com/office/50746fa8-630a-406b-81d0-4a2aed395543) | Retorna o arco tangente de um número |
| [Função ATAN2](https://support.microsoft.com/office/c04592ab-b9e3-4908-b428-c96b3a565033) | Retorna o arco tangente das coordenadas x e y |
| [Função ATANH](https://support.microsoft.com/office/3cd65768-0de7-4f1d-b312-d01c8c930d90) | Retorna a tangente hiperbólica inversa de um número |
| [Função DESV.MÉDIO](https://support.microsoft.com/office/58fe8d65-2a84-4dc7-8052-f3f87b5c6639) | Retorna a média dos desvios absolutos dos pontos de dados a partir de sua média |
| [Função MÉDIA](https://support.microsoft.com/office/047bac88-d466-426c-a32b-8f33eb960cf6) | Retorna a média dos argumentos |
| [Função MÉDIAA](https://support.microsoft.com/office/f5f84098-d453-4f4c-bbba-3d2c66356091) | Retorna a média dos argumentos, incluindo números, texto e valores lógicos |
| [Função MÉDIASE](https://support.microsoft.com/office/faec8e2e-0dec-4308-af69-f5576d8ac642) | Retorna a média (média aritmética) de todas as células em um intervalo que atendem a um determinado critério |
| [Função MÉDIASES](https://support.microsoft.com/office/48910c45-1fc0-4389-a028-f7c5c3001690) | Retorna a média (média aritmética) de todas as células que satisfazem vários critérios |
| [Função BAHTTEXT](https://support.microsoft.com/office/5ba4d0b4-abd3-4325-8d22-7a92d59aab9c) | Converte um número em texto, usando o formato de moeda ß (baht) |
| [Função BASE](https://support.microsoft.com/office/2ef61411-aee9-4f29-a811-1c42456c6342) | Converte um número em uma representação de texto com a determinada base |
| [Função BESSELI](https://support.microsoft.com/office/8d33855c-9a8d-444b-98e0-852267b1c0df) | Retorna a função de Bessel In(x) modificada |
| [Função BESSELJ](https://support.microsoft.com/office/839cb181-48de-408b-9d80-bd02982d94f7) | Retorna a função de Bessel Jn(x) |
| [Função BESSELK](https://support.microsoft.com/office/606d11bc-06d3-4d53-9ecb-2803e2b90b70) | Retorna a função de Bessel Kn(x) modificada |
| [Função BESSELY](https://support.microsoft.com/office/f3a356b3-da89-42c3-8974-2da54d6353a2) | Retorna a função de Bessel Yn(x) |
| [Função DIST.BETA](https://support.microsoft.com/office/11188c9c-780a-42c7-ba43-9ecb5a878d31) | Retorna a função de distribuição cumulativa beta |
| [Função INV.BETA](https://support.microsoft.com/office/e84cb8aa-8df0-4cf6-9892-83a341d252eb) | Retorna o inverso da função de distribuição cumulativa para uma distribuição beta especificada |
| [Função BINADEC](https://support.microsoft.com/office/63905b57-b3a0-453d-99f4-647bb519cd6c) | Converte um número binário em decimal |
| [Função BINAHEX](https://support.microsoft.com/office/0375e507-f5e5-4077-9af8-28d84f9f41cc) | Converte um número binário em hexadecimal |
| [Função BINAOCT](https://support.microsoft.com/office/0a4e01ba-ac8d-4158-9b29-16c25c4c23fd) | Converte um número binário em octal |
| [Função DISTR.BINOM](https://support.microsoft.com/office/c5ae37b6-f39c-4be2-94c2-509a1480770c) | Retorna a probabilidade de distribuição binomial do termo individual |
| [Função INTERV.DISTR.BINOM](https://support.microsoft.com/office/17331329-74c7-4053-bb4c-6653a7421595) | Retorna a probabilidade de um resultado de teste usando uma distribuição binomial |
| [Função INV.BINOM](https://support.microsoft.com/office/80a0370c-ada6-49b4-83e7-05a91ba77ac9) | Retorna o menor valor para o qual a distribuição binomial cumulativa é maior ou igual ao valor padrão |
| [Função BITAND](https://support.microsoft.com/office/8a2be3d7-91c3-4b48-9517-64548008563a) | Retorna um bit "E" de dois números |
| [Função DESLOCESQBIT](https://support.microsoft.com/office/c55bb27e-cacd-4c7c-b258-d80861a03c9c) | Retorna um valor numérico deslocado à esquerda por quantidade_deslocamento bits |
| [Função BITOR](https://support.microsoft.com/office/f6ead5c8-5b98-4c9e-9053-8ad5234919b2) | Retorna um bit "OU" de dois números |
| [Função DESLOCDIRBIT](https://support.microsoft.com/office/274d6996-f42c-4743-abdb-4ff95351222c) | Retorna um valor numérico deslocado à direita por quantidade_deslocamento bits |
| [Função BITXOR](https://support.microsoft.com/office/c81306a1-03f9-4e89-85ac-b86c3cba10e4) | Retorna um bit 'Exclusivo Ou' de dois números |
| [CEILING. MATEMÁTICA, ECMA_CEILING funções](https://support.microsoft.com/office/80f95d2f-b499-4eee-9f16-f795a8e306c8) | Arredonda um número para cima, para o inteiro mais próximo ou para o múltiplo mais próximo significativo |
| [Função TETO.PRECISO](https://support.microsoft.com/office/f366a774-527a-4c92-ba49-af0a196e66cb) | Arredonda um número para o inteiro mais próximo ou para o múltiplo mais próximo significativo. Independentemente do sinal do número, ele é arredondado para cima. |
| [Função CARACT](https://support.microsoft.com/office/bbd249c8-b36e-4a91-8017-1c133f9b837a) | Retorna o caractere especificado pelo número de código |
| [Função DIST.QUIQUA](https://support.microsoft.com/office/8486b05e-5c05-4942-a9ea-f6b341518732) | Retorna a função de densidade da probabilidade beta cumulativa |
| [Função DIST.QUIQUA.CD](https://support.microsoft.com/office/dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2) | Retorna a probabilidade unicaudal da distribuição qui-quadrada |
| [Função INV.QUIQUA](https://support.microsoft.com/office/400db556-62b3-472d-80b3-254723e7092f) | Retorna a função de densidade da probabilidade beta cumulativa |
| [Função INV.QUIQUA.CD](https://support.microsoft.com/office/435b5ed8-98d5-4da6-823f-293e2cbc94fe) | Retorna o inverso da probabilidade unicaudal da distribuição qui-quadrada |
| [Função ESCOLHER](https://support.microsoft.com/office/fc5c184f-cb62-4ec7-a46e-38653b98f5bc) | Escolhe um valor em uma lista de valores |
| [Função TIRAR](https://support.microsoft.com/office/26f3d7c5-475f-4a9c-90e5-4b8ba987ba41) | Remove do texto todos os caracteres não imprimíveis |
| [Função CÓDIGO](https://support.microsoft.com/office/c32b692b-2ed0-4a04-bdd9-75640144b928) | Retorna um código numérico para o primeiro caractere de uma cadeia de texto |
| [Função COLS](https://support.microsoft.com/office/4e8e7b4e-e603-43e8-b177-956088fa48ca) | Retorna o número de colunas em uma referência |
| [Função COMBIN](https://support.microsoft.com/office/12a3f276-0a21-423a-8de6-06990aaf638a) | Retorna o número de combinações de um determinado número de objetos |
| [Função COMBINA](https://support.microsoft.com/office/efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d) | Retorna o número de combinações com repetições de um determinado número de itens |
| [Função COMPLEXO](https://support.microsoft.com/office/f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128) | Converte coeficientes reais e imaginários em um número complexo |
| [Função CONCATENAR](https://support.microsoft.com/office/8f8ae884-2ca8-4f7a-b093-75d702bea31d) | Agrupa vários itens de texto em um item de texto |
| [Função INT.CONFIANÇA.NORM](https://support.microsoft.com/office/7cec58a6-85bb-488d-91c3-63828d4fbfd4) | Retorna o intervalo de confiança para um meio de preenchimento |
| [Função INT.CONFIANÇA.T](https://support.microsoft.com/office/e8eca395-6c3a-4ba9-9003-79ccc61d3c53) | Retorna o intervalo de confiança para um meio de preenchimento, usando a distribuição t de Student |
| [Função CONVERTER](https://support.microsoft.com/office/d785bef1-808e-4aac-bdcd-666c810f9af2) | Converte um número de um sistema de medidas para outro |
| [Função COS](https://support.microsoft.com/office/0fb808a5-95d6-4553-8148-22aebdce5f05) | Retorna o cosseno de um número |
| [Função COSH](https://support.microsoft.com/office/e460d426-c471-43e8-9540-a57ff3b70555) | Retorna o cosseno hiperbólico de um número |
| [Função COT](https://support.microsoft.com/office/c446f34d-6fe4-40dc-84f8-cf59e5f5e31a) | Retorna a cotangente de um ângulo |
| [Função COTH](https://support.microsoft.com/office/2e0b4cb6-0ba0-403e-aed4-deaa71b49df5) | Retorna a cotangente hiperbólica de um número |
| [Função CONT.NÚM](https://support.microsoft.com/office/a59cd7fc-b623-4d93-87a4-d23bf411294c) | Calcula quantos números há na lista de argumentos |
| [Função CONT.VALORES](https://support.microsoft.com/office/7dc98875-d5c1-46f1-9a82-53f3219e2509) | Calcula quantos valores há na lista de argumentos |
| [Função CONTAR.VAZIO](https://support.microsoft.com/office/6a92d772-675c-4bee-b346-24af6bd3ac22) | Conta o número de células vazias no intervalo especificado |
| [Função CONT.SE](https://support.microsoft.com/office/e0de10c6-f885-4e71-abb4-1f464816df34) | Conta o número de células em um intervalo que atendem aos critérios fornecidos |
| [Função CONT.SES](https://support.microsoft.com/office/dda3dc6e-f74e-4aee-88bc-aa8c2a866842) | Conta o número de células dentro de um intervalo que atende a múltiplos critérios |
| [Função CUPDIASINLIQ](https://support.microsoft.com/office/eb9a8dfb-2fb2-4c61-8e5d-690b320cf872) | Retorna o número de dias do início do período de cupom até a data de liquidação |
| [Função CUPDIAS](https://support.microsoft.com/office/cc64380b-315b-4e7b-950c-b30b0a76f671) | Retorna o número de dias no período de cupom que contém a data de liquidação |
| [Função CUPDIASPRÓX](https://support.microsoft.com/office/5ab3f0b2-029f-4a8b-bb65-47d525eea547) | Retorna o número de dias da data de liquidação até a data do próximo cupom |
| [Função CUPDATAPRÓX](https://support.microsoft.com/office/fd962fef-506b-4d9d-8590-16df5393691f) | Retorna a próxima data de cupom após a data de quitação |
| [Função CUPNÚM](https://support.microsoft.com/office/a90af57b-de53-4969-9c99-dd6139db2522) | Retorna o número de cupons pagáveis entre as datas de quitação e vencimento |
| [Função CUPDATAANT](https://support.microsoft.com/office/2eb50473-6ee9-4052-a206-77a9a385d5b3) | Retorna a data de cupom anterior à data de quitação |
| [Função COSEC](https://support.microsoft.com/office/07379361-219a-4398-8675-07ddc4f135c1) | Retorna a cossecante de um ângulo |
| [Função COSECH](https://support.microsoft.com/office/f58f2c22-eb75-4dd6-84f4-a503527f8eeb) | Retorna a cossecante hiperbólica de um ângulo |
| [Função PGTOJURACUM](https://support.microsoft.com/office/61067bb0-9016-427d-b95b-1a752af0e606) | Retorna os juros acumulados pagos entre dois períodos |
| [Função PGTOCAPACUM](https://support.microsoft.com/office/94a4516d-bd65-41a1-bc16-053a6af4c04d) | Retorna o capital acumulado pago sobre um empréstimo entre dois períodos |
| [Função DATA](https://support.microsoft.com/office/e36c0c8c-4104-49da-ab83-82328b832349) | Retorna o número de série de uma data específica |
| [Função DATA.VALOR](https://support.microsoft.com/office/df8b07d4-7761-4a93-bc33-b7471bbff252) | Converte uma data na forma de texto em um número de série |
| [Função BDMÉDIA](https://support.microsoft.com/office/a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee) | Retorna a média das entradas selecionadas de um banco de dados |
| [Função DIA](https://support.microsoft.com/office/8a7d1cbb-6c7d-4ba1-8aea-25c134d03101) | Converte um número de série em um dia do mês |
| [Função DIAS](https://support.microsoft.com/office/57740535-d549-4395-8728-0f07bff0b9df) | Retorna o número de dias entre duas datas |
| [Função DIAS360](https://support.microsoft.com/office/b9a509fd-49ef-407e-94df-0cbda5718c2a) | Calcula o número de dias entre duas datas com base em um ano de 360 dias |
| [Função BD](https://support.microsoft.com/office/354e7d28-5f93-4ff1-8a52-eb4ee549d9d7) | Retorna a depreciação de um ativo para um período especificado, usando o método de balanço de declínio fixo |
| [Função DBCS](https://support.microsoft.com/office/a4025e73-63d2-4958-9423-21a24794c9e5) | Altera letras do inglês ou katakana de meia largura (byte único) dentro de uma cadeia de caracteres para caracteres de largura total (bytes duplos) |
| [Função BDCONTAR](https://support.microsoft.com/office/c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1) | Conta as células que contêm números em um banco de dados |
| [Função BDCONTARA](https://support.microsoft.com/office/00232a6d-5a66-4a01-a25b-c1653fda1244) | Conta células não vazias em um banco de dados |
| [Função BDD](https://support.microsoft.com/office/519a7a37-8772-4c96-85c0-ed2c209717a5) | Retorna a depreciação de um ativo com relação a um período especificado usando o método de saldos decrescentes duplos ou qualquer outro método especificado por você |
| [Função DECABIN](https://support.microsoft.com/office/0f63dd0e-5d1a-42d8-b511-5bf5c6d43838) | Converte um número decimal em binário |
| [Função DECAHEX](https://support.microsoft.com/office/6344ee8b-b6b5-4c6a-a672-f64666704619) | Converte um número decimal em hexadecimal |
| [Função DECAOCT](https://support.microsoft.com/office/c9d835ca-20b7-40c4-8a9e-d3be351ce00f) | Converte um número decimal em octal |
| [Função DECIMAL](https://support.microsoft.com/office/ee554665-6176-46ef-82de-0a283658da2e) | Converte em um número decimal a representação de texto de um número em determinada base |
| [Função GRAUS](https://support.microsoft.com/office/4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1) | Converte radianos em graus |
| [Função DELTA](https://support.microsoft.com/office/2f763672-c959-4e07-ac33-fe03220ba432) | Testa se dois valores são iguais |
| [Função DESVQ](https://support.microsoft.com/office/8b739616-8376-4df5-8bd0-cfe0a6caf444) | Retorna a soma dos quadrados dos desvios |
| [Função BDEXTRAIR](https://support.microsoft.com/office/455568bf-4eef-45f7-90f0-ec250d00892e) | Extrai de um banco de dados um único registro que corresponde aos critérios especificados |
| [Função DESC](https://support.microsoft.com/office/71fce9f3-3f05-4acf-a5a3-eac6ef4daa53) | Retorna a taxa de desconto de um título |
| [Função BDMÁX](https://support.microsoft.com/office/f4e8209d-8958-4c3d-a1ee-6351665d41c2) | Retorna o valor máximo de entradas selecionadas de banco de dados |
| [Função BDMÍN](https://support.microsoft.com/office/4ae6f1d9-1f26-40f1-a783-6dc3680192a3) | Retorna o valor mínimo de entradas selecionadas de um banco de dados |
| [Funções DOLLAR, USDOLLAR](https://support.microsoft.com/office/a6cd05d9-9740-4ad3-a469-8109d18ff611) | Converte um número em texto, usando o formato de moeda $ (cifrão) |
| [Função MOEDADEC](https://support.microsoft.com/office/db85aab0-1677-428a-9dfd-a38476693427) | Converte um preço em moeda expresso como uma fração em um preço em moeda expresso como um número decimal |
| [Função MOEDAFRA](https://support.microsoft.com/office/0835d163-3023-4a33-9824-3042c5d4f495) | Converte um preço em moeda expresso como um número decimal em um preço em moeda expresso como uma fração |
| [Função BDMULTIPL](https://support.microsoft.com/office/4f96b13e-d49c-47a7-b769-22f6d017cb31) | Multiplica os valores em um campo específico de registros que correspondem ao critério em um banco de dados |
| [Função BDEST](https://support.microsoft.com/office/026b8c73-616d-4b5e-b072-241871c4ab96) | Estima o desvio padrão com base em uma amostra de entradas selecionadas de um banco de dados |
| [Função BDDESVPA](https://support.microsoft.com/office/04b78995-da03-4813-bbd9-d74fd0f5d94b) | Calcula o desvio padrão com base no preenchimento completo de entradas selecionadas de banco de dados |
| [Função BDSOMA](https://support.microsoft.com/office/53181285-0c4b-4f5a-aaa3-529a322be41b) | Soma os números na coluna de campos de registros do banco de dados que correspondem ao critério |
| [Função DURAÇÃO](https://support.microsoft.com/office/b254ea57-eadc-4602-a86a-c8e369334038) | Retorna a duração anual de um título com pagamentos de juros periódicos |
| [Função Dlet](https://support.microsoft.com/office/d6747ca9-99c7-48bb-996e-9d7af00f3ed1) | Estima a variação com base em uma amostra de entradas selecionadas de um banco de dados |
| [Função BDVARP](https://support.microsoft.com/office/eb0ba387-9cb7-45c8-81e9-0394912502fc) | Calcula a variação com base no preenchimento completo de entradas selecionadas de um banco de dados |
| [Função DATAM](https://support.microsoft.com/office/3c920eb2-6e66-44e7-a1f5-753ae47ee4f5) | Retorna o número de série da data que é o número indicado de meses antes ou depois da data inicial |
| [Função EFETIVA](https://support.microsoft.com/office/910d4e4c-79e2-4009-95e6-507e04f11bc4) | Retorna a taxa de juros anual efetiva |
| [Função FIMMÊS](https://support.microsoft.com/office/7314ffa1-2bc9-4005-9d66-f49db127d628) | Retorna o número de série do último dia do mês antes ou depois de um número especificado de meses |
| [Função FUNERRO](https://support.microsoft.com/office/c53c7e7b-5482-4b6c-883e-56df3c9af349) | Retorna a função de erro |
| [Função FUNERRO.PRECISO](https://support.microsoft.com/office/9a349593-705c-4278-9a98-e4122831a8e0) | Retorna a função de erro |
| [Função FUNERROCOMPL](https://support.microsoft.com/office/736e0318-70ba-4e8b-8d08-461fe68b71b3) | Retorna a função de erro complementar |
| [Função FUNERROCOMPL.PRECISO](https://support.microsoft.com/office/e90e6bab-f45e-45df-b2ac-cd2eb4d4a273) | Retorna a função FUNERRO complementar integrada entre x e infinito |
| [Função TIPO.ERRO](https://support.microsoft.com/office/10958677-7c8d-44f7-ae77-b9a9ee6eefaa) | Retorna um número correspondente a um tipo de erro |
| [Função PAR](https://support.microsoft.com/office/197b5f06-c795-4c1e-8696-3c3b8a646cf9) | Arredonda um número para cima até o inteiro par mais próximo |
| [Função EXATO](https://support.microsoft.com/office/d3087698-fc15-4a15-9631-12575cf29926) | Verifica se dois valores de texto são idênticos |
| [Função EXP](https://support.microsoft.com/office/c578f034-2c45-4c37-bc8c-329660a63abe) | Retorna e elevado à potência de um número especificado |
| [Função DISTR.EXPON](https://support.microsoft.com/office/4c12ae24-e563-4155-bf3e-8b78b6ae140e) | Retorna a distribuição exponencial |
| [Função DIST.F](https://support.microsoft.com/office/a887efdc-7c8e-46cb-a74a-f884cd29b25d) | Retorna a distribuição de probabilidade F |
| [Função DIST.F.CD](https://support.microsoft.com/office/d74cbb00-6017-4ac9-b7d7-6049badc0520) | Retorna a distribuição de probabilidade F |
| [Função INV.F](https://support.microsoft.com/office/0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe) | Retorna o inverso da distribuição de probabilidade F |
| [Função INV.F.CD](https://support.microsoft.com/office/d371aa8f-b0b1-40ef-9cc2-496f0693ac00) | Retorna o inverso da distribuição de probabilidade F |
| [Função FATORIAL](https://support.microsoft.com/office/ca8588c2-15f2-41c0-8e8c-c11bd471a4f3) | Retorna o fatorial de um número |
| [Função FATDUPLO](https://support.microsoft.com/office/e67697ac-d214-48eb-b7b7-cce2589ecac8) | Retorna o fatorial duplo de um número |
| [Função FALSO](https://support.microsoft.com/office/2d58dfa5-9c03-4259-bf8f-f0ae14346904) | Retorna o valor lógico `FALSE` |
| [Funções PROCURAR, PROCURARB](https://support.microsoft.com/office/c7912941-af2a-4bdf-a553-d0d89b0a0628) | Procura um valor de texto dentro de outro (diferencia maiúsculas de minúsculas) |
| [Função FISHER](https://support.microsoft.com/office/d656523c-5076-4f95-b87b-7741bf236c69) | Retorna a transformação Fisher |
| [Função FISHERINV](https://support.microsoft.com/office/62504b39-415a-4284-a285-19c8e82f86bb) | Retorna o inverso da transformação Fisher |
| [Função FIXO](https://support.microsoft.com/office/ffd5723c-324c-45e9-8b96-e41be2a8274a) | Formata um número como texto com um número fixo de decimais |
| [Função de ARREDMULTB.MAT](https://support.microsoft.com/office/c302b599-fbdb-4177-ba19-2c2b1249a2f5) | Arredonda um número para baixo para o inteiro mais próximo ou para o múltiplo mais próximo de significância |
| [Função ARREDMULTB.PRECISO](https://support.microsoft.com/office/f769b468-1452-4617-8dc3-02f842a0702e) | Arredonda um número para baixo para o inteiro mais próximo ou para o múltiplo mais próximo de significância. Independentemente do sinal do número, ele é arredondado para baixo. |
| [Função VF](https://support.microsoft.com/office/2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3) | Retorna o valor futuro de um investimento |
| [Função VFPLANO](https://support.microsoft.com/office/bec29522-bd87-4082-bab9-a241f3fb251d) | Retorna o valor futuro de um capital inicial após a aplicação de uma série de taxas de juros compostas |
| [Função GAMA](https://support.microsoft.com/office/ce1702b1-cf55-471d-8307-f83be0fc5297) | Retorna o valor da função GAMA |
| [Função DIST.GAMA](https://support.microsoft.com/office/9b6f1538-d11c-4d5f-8966-21f6a2201def) | Retorna a distribuição gama |
| [Função INV.GAMA](https://support.microsoft.com/office/74991443-c2b0-4be5-aaab-1aa4d71fbb18) | Retorna o inverso da distribuição cumulativa gama |
| [Função LNGAMA](https://support.microsoft.com/office/b838c48b-c65f-484f-9e1d-141c55470eb9) | Retorna o logaritmo natural da função gama, G(x) |
| [Função LNGAMA.PRECISO](https://support.microsoft.com/office/5cdfe601-4e1e-4189-9d74-241ef1caa599) | Retorna o logaritmo natural da função gama, G(x) |
| [Função GAUSS](https://support.microsoft.com/office/069f1b4e-7dee-4d6a-a71f-4b69044a6b33) | Retorna menos 0,5 que a distribuição cumulativa normal padrão |
| [Função MDC](https://support.microsoft.com/office/d5107a51-69e3-461f-8e4c-ddfc21b5073a) | Retorna o máximo divisor comum |
| [Função MÉDIA.GEOMÉTRICA](https://support.microsoft.com/office/db1ac48d-25a5-40a0-ab83-0b38980e40d5) | Retorna a média geométrica |
| [Função DEGRAU](https://support.microsoft.com/office/f37e7d2a-41da-4129-be95-640883fca9df) | Testa se um número é maior do que um valor limite |
| [Função MÉDIA.HARMÔNICA](https://support.microsoft.com/office/5efd9184-fab5-42f9-b1d3-57883a1d3bc6) | Retorna a média harmônica |
| [Função HEXABIN](https://support.microsoft.com/office/a13aafaa-5737-4920-8424-643e581828c1) | Converte um número hexadecimal em binário |
| [Função HEXADEC](https://support.microsoft.com/office/8c8c3155-9f37-45a5-a3ee-ee5379ef106e) | Converte um número hexadecimal em decimal |
| [Função HEXAOCT](https://support.microsoft.com/office/54d52808-5d19-4bd0-8a63-1096a5d11912) | Converte um número hexadecimal em octal |
| [Função PROCH](https://support.microsoft.com/office/a3034eec-b719-4ba3-bb65-e1ad662ed95f) | Procura na linha superior de uma matriz e retorna o valor da célula especificada |
| [Função HORA](https://support.microsoft.com/office/a3afa879-86cb-4339-b1b5-2dd2d7310ac7) | Converte um número de série em um hora |
| [Função HIPERLINK](https://support.microsoft.com/office/333c7ce6-c5ae-4164-9c47-7de9b76f577f) | Cria um atalho ou salto que abre um documento armazenado em um servidor de rede, uma intranet ou na Internet |
| [Função DIST.HIPERGEOM.N](https://support.microsoft.com/office/6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf) | Retorna a distribuição hipergeométrica |
| [Função SE](https://support.microsoft.com/office/69aed7c9-4e8a-4755-a9bc-aa8bbff73be2) | Especifica um teste lógico a ser executado |
| [Função IMABS](https://support.microsoft.com/office/b31e73c6-d90c-4062-90bc-8eb351d765a1) | Retorna o valor absoluto (módulo) de um número complexo |
| [Função IMAGINÁRIO](https://support.microsoft.com/office/dd5952fd-473d-44d9-95a1-9a17b23e428a) | Retorna o coeficiente imaginário de um número complexo |
| [Função IMARG](https://support.microsoft.com/office/eed37ec1-23b3-4f59-b9f3-d340358a034a) | Retorna o argumento teta, um ângulo expresso em radianos |
| [Função IMCONJ](https://support.microsoft.com/office/2e2fc1ea-f32b-4f9b-9de6-233853bafd42) | Retorna o conjugado complexo de um número complexo |
| [Função IMCOS](https://support.microsoft.com/office/dad75277-f592-4a6b-ad6c-be93a808a53c) | Retorna o cosseno de um número complexo |
| [Função IMCOSH](https://support.microsoft.com/office/053e4ddb-4122-458b-be9a-457c405e90ff) | Retorna o cosseno hiperbólico de um número complexo |
| [Função IMCOT](https://support.microsoft.com/office/dc6a3607-d26a-4d06-8b41-8931da36442c) | Retorna a cotangente de um número complexo |
| [Função IMCOSEC](https://support.microsoft.com/office/9e158d8f-2ddf-46cd-9b1d-98e29904a323) | Retorna a cossecante de um número complexo |
| [Função IMCOSECH](https://support.microsoft.com/office/c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9) | Retorna a cossecante hiperbólica de um número complexo |
| [Função IMDIV](https://support.microsoft.com/office/a505aff7-af8a-4451-8142-77ec3d74d83f) | Retorna o quociente de dois números complexos |
| [Função IMEXP](https://support.microsoft.com/office/c6f8da1f-e024-4c0c-b802-a60e7147a95f) | Retorna o exponencial de um número complexo |
| [Função IMLN](https://support.microsoft.com/office/32b98bcf-8b81-437c-a636-6fb3aad509d8) | Retorna o logaritmo natural de um número complexo |
| [Função IMLOG10](https://support.microsoft.com/office/58200fca-e2a2-4271-8a98-ccd4360213a5) | Retorna o logaritmo de base 10 de um número complexo |
| [Função IMLOG2](https://support.microsoft.com/office/152e13b4-bc79-486c-a243-e6a676878c51) | Retorna o logaritmo de base 2 de um número complexo |
| [Função IMPOT](https://support.microsoft.com/office/210fd2f5-f8ff-4c6a-9d60-30e34fbdef39) | Retorna um número complexo elevado a uma potência inteira |
| [Função IMPROD](https://support.microsoft.com/office/2fb8651a-a4f2-444f-975e-8ba7aab3a5ba) | Retorna o produto de 2 a 255 números complexos |
| [Função IMREAL](https://support.microsoft.com/office/d12bc4c0-25d0-4bb3-a25f-ece1938bf366) | Retorna o coeficiente real de um número complexo |
| [Função IMSEC](https://support.microsoft.com/office/6df11132-4411-4df4-a3dc-1f17372459e0) | Retorna a secante de um número complexo |
| [Função IMSECH](https://support.microsoft.com/office/f250304f-788b-4505-954e-eb01fa50903b) | Retorna a secante hiperbólica de um número complexo |
| [Função IMSENO](https://support.microsoft.com/office/1ab02a39-a721-48de-82ef-f52bf37859f6) | Retorna o seno de um número complexo |
| [Função IMSENH](https://support.microsoft.com/office/dfb9ec9e-8783-4985-8c42-b028e9e8da3d) | Retorna o seno hiperbólico de um número complexo |
| [Função IMRAIZ](https://support.microsoft.com/office/e1753f80-ba11-4664-a10e-e17368396b70) | Retorna a raiz quadrada de um número complexo |
| [Função IMSUBTR](https://support.microsoft.com/office/2e404b4d-4935-4e85-9f52-cb08b9a45054) | Retorna a diferença entre dois números complexos |
| [Função IMSOMA](https://support.microsoft.com/office/81542999-5f1c-4da6-9ffe-f1d7aaa9457f) | Retorna a soma de números complexos |
| [Função IMTAN](https://support.microsoft.com/office/8478f45d-610a-43cf-8544-9fc0b553a132) | Retorna a tangente de um número complexo |
| [Função INT](https://support.microsoft.com/office/a6c4af9e-356d-4369-ab6a-cb1fd9d343ef) | Arredonda um número para baixo até o número inteiro mais próximo |
| [Função TAXAJUROS](https://support.microsoft.com/office/5cb34dde-a221-4cb6-b3eb-0b9e55e1316f) | Retorna a taxa de juros de um título totalmente investido |
| [Função IPGTO](https://support.microsoft.com/office/5cce0ad6-8402-4a41-8d29-61a0b054cb6f) | Retorna o pagamento de juros para um investimento em um determinado período |
| [Função TIR](https://support.microsoft.com/office/64925eaa-9988-495b-b290-3ad0c163c1bc) | Retorna a taxa interna de retorno de uma série de fluxos de caixa |
| [Função ÉERRO](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for um valor de erro diferente de #N/D |
| [Função ÉERROS](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for um valor de erro |
| [Função ÉPAR](https://support.microsoft.com/office/aa15929a-d77b-4fbb-92f4-2f479af55356) | Retorna `TRUE` se o número for par |
| [Função ÉFÓRMULA](https://support.microsoft.com/office/e4d1355f-7121-4ef2-801e-3839bfd6b1e5) | Retorna `TRUE` quando há uma referência a uma célula que contém uma fórmula |
| [Função ÉLÓGICO](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for um valor lógico |
| [Função É.NÃO.DISP](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for o valor de erro #N/D |
| [Função É.NÃO.TEXTO](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for diferente de texto |
| [Função ÉNÚM](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for um número |
| [Função ISO.TETO](https://support.microsoft.com/office/e587bb73-6cc2-4113-b664-ff5b09859a83) | Retorna um número para o inteiro mais próximo ou para o múltiplo mais próximo de significância |
| [Função ÉIMPAR](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o número for ímpar |
| [Função NÚMSEMANAISO](https://support.microsoft.com/office/1c2d0afe-d25b-4ab1-8894-8d0520e90e0e) | Retorna o número do número da semana ISO do ano referente a determinada data |
| [Função ÉPGTO](https://support.microsoft.com/office/fa58adb6-9d39-4ce0-8f43-75399cea56cc) | Calcula os juros pagos durante um período específico de um investimento |
| [Função ÉREF](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for uma referência |
| [Função ÉTEXTO](https://support.microsoft.com/office/0f2d7971-6019-40a0-a171-f2d869135665) | Retorna `TRUE` se o valor for texto |
| [Função CURT](https://support.microsoft.com/office/bc3a265c-5da4-4dcb-b7fd-c237789095ab) | Retorna a curtose de um conjunto de dados |
| [Função MAIOR](https://support.microsoft.com/office/3af0af19-1190-42bb-bb8b-01672ec00a64) | Retorna o maior valor k-ésimo em um conjunto de dados |
| [Função MMC](https://support.microsoft.com/office/7152b67a-8bb5-4075-ae5c-06ede5563c94) | Retorna o mínimo múltiplo comum |
| [Funções ESQUERDA, ESQUERDAB](https://support.microsoft.com/office/9203d2d2-7960-479b-84c6-1ea52b99640c) | Retorna os caracteres mais à esquerda de um valor de texto |
| [Funções NÚM.CARACT, NÚM.CARACTB](https://support.microsoft.com/office/29236f94-cedc-429d-affd-b5e33d2c67cb) | Retorna o número de caracteres em uma cadeia de texto |
| [Função LN](https://support.microsoft.com/office/81fe1ed7-dac9-4acd-ba1d-07a142c6118f) | Retorna o logaritmo natural de um número |
| [Função LOG](https://support.microsoft.com/office/4e82f196-1ca9-4747-8fb0-6c4a3abb3280) | Retorna o logaritmo de um número de uma base especificada |
| [Função LOG10](https://support.microsoft.com/office/c75b881b-49dd-44fb-b6f4-37e3486a0211) | Retorna o logaritmo de base 10 de um número |
| [Função DIST.LOGNORMAL.N](https://support.microsoft.com/office/eb60d00b-48a9-4217-be2b-6074aee6b070) | Retorna a distribuição lognormal cumulativa |
| [Função INV.LOGNORMAL](https://support.microsoft.com/office/fe79751a-f1f2-4af8-a0a1-e151b2d4f600) | Retorna o inverso da distribuição cumulativa lognormal |
| [Função PROC](https://support.microsoft.com/office/446d94af-663b-451d-8251-369d5e3864cb) | Procura valores em um vetor ou uma matriz |
| [Função MINÚSCULA](https://support.microsoft.com/office/3f21df02-a80c-44b2-afaf-81358f9fdeb4) | Converte texto em minúsculas |
| [Função CORRESP](https://support.microsoft.com/office/e8dffd45-c762-47d6-bf89-533f4a37673a) | Procura valores em uma referência ou matriz |
| [Função MÁXIMO](https://support.microsoft.com/office/e0012414-9ac8-4b34-9a47-73e662c08098) | Retorna o valor máximo em uma lista de argumentos |
| [Função MÁXIMOA](https://support.microsoft.com/office/814bda1e-3840-4bff-9365-2f59ac2ee62d) | Retorna o maior valor em uma lista de argumentos, incluindo números, texto e valores lógicos |
| [Função MDURAÇÃO](https://support.microsoft.com/office/b3786a69-4f20-469a-94ad-33e5b90a763c) | Retorna a duração modificada Macauley de um título com um valor de paridade equivalente a R$ 100 |
| [Função MED](https://support.microsoft.com/office/d0916313-4753-414c-8537-ce85bdd967d2) | Retorna a mediana dos números indicados |
| [Funções EXT.TEXTO, EXT.TEXTOB](https://support.microsoft.com/office/d5f9e25c-d7d6-472e-b568-4ecb12433028) | Retorna um número específico de caracteres de uma cadeia de texto começando na posição especificada |
| [Função MÍNIMO](https://support.microsoft.com/office/61635d12-920f-4ce2-a70f-96f202dcc152) | Retorna o valor mínimo em uma lista de argumentos |
| [Função MÍNIMOA](https://support.microsoft.com/office/245a6f46-7ca5-4dc7-ab49-805341bc31d3) | Retorna o menor valor em uma lista de argumentos, incluindo números, texto e valores lógicos |
| [Função MINUTO](https://support.microsoft.com/office/af728df0-05c4-4b07-9eed-a84801a60589) | Converte um número de série em um minuto |
| [Função MTIR](https://support.microsoft.com/office/b020f038-7492-4fb4-93c1-35c345b53524) | Calcula a taxa interna de retorno em que fluxos de caixa positivos e negativos são financiados com diferentes taxas |
| [Função MOD](https://support.microsoft.com/office/9b6cd169-b6ee-406a-a97b-edf2a9dc24f3) | Retorna o resto da divisão |
| [Função MÊS](https://support.microsoft.com/office/579a2881-199b-48b2-ab90-ddba0eba86e8) | Converte um número de série em um mês |
| [Função MARRED](https://support.microsoft.com/office/c299c3b0-15a5-426d-aa4b-d2d5b3baf427) | Retorna um número arredondado ao múltiplo desejado |
| [Função MULTINOMIAL](https://support.microsoft.com/office/6fa6373c-6533-41a2-a45e-a56db1db1bf6) | Retorna o multinômio de um conjunto de números |
| [Função N](https://support.microsoft.com/office/a624cad1-3635-4208-b54a-29733d1278c9) | Retorna um valor convertido em um número |
| [Função NÃO.DISP](https://support.microsoft.com/office/5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c) | Retorna o valor de erro #N/D |
| [Função DIST.BIN.NEG.N](https://support.microsoft.com/office/c8239f89-c2d0-45bd-b6af-172e570f8599) | Retorna a distribuição binomial negativa |
| [Função DIATRABALHOTOTAL](https://support.microsoft.com/office/48e717bf-a7a3-495f-969e-5005e3eb18e7) | Retorna o número de dias úteis inteiros entre duas datas |
| [Função DIATRABALHOTOTAL.INTL](https://support.microsoft.com/office/a9b26239-4f20-46a1-9ab8-4e925bfd5e28) | Retorna o número de dias de trabalho totais entre duas datas usando parâmetros para indicar quais e quantos dias caem em finais de semana |
| [Função NOMINAL](https://support.microsoft.com/office/7f1ae29b-6b92-435e-b950-ad8b190ddd2b) | Retorna a taxa de juros nominal anual |
| [Função DIST.NORM.N](https://support.microsoft.com/office/edb1cc14-a21c-4e53-839d-8082074c9f8d) | Retorna a distribuição cumulativa normal |
| [Função INV.NORM.N](https://support.microsoft.com/office/54b30935-fee7-493c-bedb-2278a9db7e13) | Retorna o inverso da distribuição cumulativa normal |
| [Função DIST.NORMP.N](https://support.microsoft.com/office/1e787282-3832-4520-a9ae-bd2a8d99ba88) | Retorna a distribuição cumulativa normal padrão |
| [Função INV.NORMP.N](https://support.microsoft.com/office/d6d556b4-ab7f-49cd-b526-5a20918452b1) | Retorna o inverso da distribuição cumulativa normal padrão |
| [Função NÃO](https://support.microsoft.com/office/9cfc6011-a054-40c7-a140-cd4ba2d87d77) | Inverte o valor lógico do argumento |
| [Função AGORA](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46) | Retorna o número de série sequencial da data e hora atuais |
| [Função NPER](https://support.microsoft.com/office/240535b5-6653-4d2d-bfcf-b6a38151d815) | Retorna o número de períodos de um investimento |
| [Função VPL](https://support.microsoft.com/office/8672cb67-2576-4d07-b67b-ac28acf2a568) | Retorna o valor líquido atual de um investimento com base em uma série de fluxos de caixa periódicos e em uma taxa de desconto |
| [Função VALORNUMÉRICO](https://support.microsoft.com/office/1b05c8cf-2bfa-4437-af70-596c7ea7d879) | Converte texto em número de maneira independente de localidade |
| [Função OCTABIN](https://support.microsoft.com/office/55383471-3c56-4d27-9522-1a8ec646c589) | Converte um número octal em binário |
| [Função OCTADEC](https://support.microsoft.com/office/87606014-cb98-44b2-8dbb-e48f8ced1554) | Converte um número octal em decimal |
| [Função OCTAHEX](https://support.microsoft.com/office/912175b4-d497-41b4-a029-221f051b858f) | Converte um número octal em hexadecimal |
| [Função ÍMPAR](https://support.microsoft.com/office/deae64eb-e08a-4c88-8b40-6d0b42575c98) | Arredonda um número para cima até o inteiro ímpar mais próximo |
| [Função PREÇOPRIMINC](https://support.microsoft.com/office/d7d664a8-34df-4233-8d2b-922bcf6a69e1) | Retorna o preço por R$ 100 do valor nominal de um título com um período inicial incompleto |
| [Função LUCROPRIMINC](https://support.microsoft.com/office/66bc8b7b-6501-4c93-9ce3-2fd16220fe37) | Retorna o rendimento de um título com um período inicial incompleto |
| [Função PREÇOÚLTINC](https://support.microsoft.com/office/fb657749-d200-4902-afaf-ed5445027fc4) | Retorna o preço por R$ 100 do valor nominal de um título com um período final incompleto |
| [Função LUCROÚLTINC](https://support.microsoft.com/office/c873d088-cf40-435f-8d41-c8232fee9238) | Retorna o rendimento de um título com um período final incompleto |
| [Função OU](https://support.microsoft.com/office/7d17ad14-8700-4281-b308-00b131e22af0) | Retorna `TRUE` se um dos argumentos for verdadeiro |
| [Função DURAÇÃOP](https://support.microsoft.com/office/44f33460-5be5-4c90-b857-22308892adaf) | Retorna o número de períodos necessários para que um investimento atinja um valor específico |
| [Função PERCENTIL.EXC](https://support.microsoft.com/office/bbaa7204-e9e1-4010-85bf-c31dc5dce4ba) | Retorna o k-ésimo percentil de valores em um intervalo, onde k está no intervalo 0..1, exclusive |
| [Função PERCENTIL.INC](https://support.microsoft.com/office/680f9539-45eb-410b-9a5e-c1355e5fe2ed) | Retorna o k-ésimo percentil de valores em um intervalo |
| [Função ORDEM.PORCENTUAL.EXC](https://support.microsoft.com/office/d8afee96-b7e2-4a2f-8c01-8fcdedaa6314) | Retorna a posição de um valor em um conjunto de dados como uma porcentagem (0..1, exclusivo) do conjunto de dados |
| [Função ORDEM.PORCENTUAL.INC](https://support.microsoft.com/office/149592c9-00c0-49ba-86c1-c1f45b80463a) | Retorna a ordem percentual de um valor em um conjunto de dados |
| [Função PERMUT](https://support.microsoft.com/office/3bd1cb9a-2880-41ab-a197-f246a7a602d3) | Retorna o número de permutações de um determinado número de objetos |
| [Função PERMUTAS](https://support.microsoft.com/office/6c7d7fdc-d657-44e6-aa19-2857b25cae4e) | Retorna o número de permutações referentes a determinado número de objetos (com repetições) que podem ser selecionadas do total de objetos |
| [Função PHI](https://support.microsoft.com/office/23e49bc6-a8e8-402d-98d3-9ded87f6295c) | Retorna o valor da função de densidade referente a uma distribuição normal padrão |
| [Função PI](https://support.microsoft.com/office/264199d0-a3ba-46b8-975a-c4a04608989b) | Retorna o valor de pi |
| [Função PGTO](https://support.microsoft.com/office/0214da64-9a63-4996-bc20-214433fa6441) | Retorna o pagamento periódico de uma anuidade |
| [Função DIST.POISSON](https://support.microsoft.com/office/8fe148ff-39a2-46cb-abf3-7772695d9636) | Retorna a distribuição de Poisson |
| [Função POTÊNCIA](https://support.microsoft.com/office/d3f2908b-56f4-4c3f-895a-07fb519c362a) | Retorna o resultado de um número elevado a uma potência |
| [Função PPGTO](https://support.microsoft.com/office/c370d9e3-7749-4ca4-beea-b06c6ac95e1b) | Retorna o pagamento de capital para determinado período de investimento |
| [Função PREÇO](https://support.microsoft.com/office/3ea9deac-8dfa-436f-a7c8-17ea02c21b0a) | Retorna o preço pelo valor nominal R$100 de um título que paga juros periódicos |
| [Função PREÇODESC](https://support.microsoft.com/office/d06ad7c1-380e-4be7-9fd9-75e3079acfd3) | Retorna o preço por valor nominal de R$ 100,00 de um título descontado |
| [Função PREÇOVENC](https://support.microsoft.com/office/52c3b4da-bc7e-476a-989f-a95f675cae77) | Retorna o preço pelo valor nominal R$100 de um título que paga juros no vencimento |
| [Função MULT](https://support.microsoft.com/office/8e6b5b24-90ee-4650-aeec-80982a0512ce) | Multiplica seus argumentos |
| [Função PRI.MAIÚSCULA](https://support.microsoft.com/office/52a5a283-e8b2-49be-8506-b2887b889f94) | Coloca a primeira letra de cada palavra em maiúscula em um valor de texto |
| [Função VP](https://support.microsoft.com/office/23879d31-0e02-4321-be01-da16e8168cbd) | Retorna o valor presente de um investimento |
| [Função QUARTIL.EXC](https://support.microsoft.com/office/5a355b7a-840b-4a01-b0f1-f538c2864cad) | Retorna o quartil do conjunto de dados, com base em valores de percentil de 0..1, exclusive |
| [Função QUARTIL.INC](https://support.microsoft.com/office/1bbacc80-5075-42f1-aed6-47d735c4819d) | Retorna o quartil de um conjunto de dados |
| [Função QUOCIENTE](https://support.microsoft.com/office/9f7bf099-2a18-4282-8fa4-65290cc99dee) | Retorna a parte inteira de uma divisão |
| [Função RADIANOS](https://support.microsoft.com/office/ac409508-3d48-45f5-ac02-1497c92de5bf) | Converte graus em radianos |
| [Função ALEATÓRIO](https://support.microsoft.com/office/4cbfa695-8869-4788-8d90-021ea9f5be73) | Retorna um número aleatório entre 0 e 1 |
| [Função ALEATÓRIOENTRE](https://support.microsoft.com/office/4cc7f0d1-87dc-4eb7-987f-a469ab381685) | Retorna um número aleatório entre os números especificados |
| [Função ORDEM.MÉD](https://support.microsoft.com/office/bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a) | Retorna a posição de um número em uma lista de números |
| [Função ORDEM.EQ](https://support.microsoft.com/office/284858ce-8ef6-450e-b662-26245be04a40) | Retorna a posição de um número em uma lista de números |
| [Função TAXA](https://support.microsoft.com/office/9f665657-4a7e-4bb7-a030-83fc59e748ce) | Retorna a taxa de juros por período de uma anuidade |
| [Função RECEBIDO](https://support.microsoft.com/office/7a3f8b93-6611-4f81-8576-828312c9b5e5) | Retorna a quantia recebida no vencimento de um título totalmente investido |
| [Funções MUDAR, SUBSTITUIRB](https://support.microsoft.com/office/8d799074-2425-4a8a-84bc-82472868878a) | Muda os caracteres dentro do texto |
| [Função REPT](https://support.microsoft.com/office/04c4d778-e712-43b4-9c15-d656582bb061) | Repete o texto um determinado número de vezes |
| [Funções DIREITA, DIREITAB](https://support.microsoft.com/office/240267ee-9afa-4639-a02b-f19e1786cf2f) | Retorna os caracteres mais à direita de um valor de texto |
| [Função ROMANO](https://support.microsoft.com/office/d6b0b99e-de46-4704-a518-b45a0f8b56f5) | Converte um algarismo arábico em romano, como texto |
| [Função ARRED](https://support.microsoft.com/office/c018c5d8-40fb-4053-90b1-b3e7f61a213c) | Arredonda um número até uma quantidade especificada de dígitos |
| [Função ARREDONDAR.PARA.BAIXO](https://support.microsoft.com/office/2ec94c73-241f-4b01-8c6f-17e6d7968f53) | Arredonda um número para baixo até zero |
| [Função ARREDONDAR.PARA.CIMA](https://support.microsoft.com/office/f8bc9b23-e795-47db-8703-db171d0c42a7) | Arredonda um número para cima afastando-o de zero |
| [Função LINS](https://support.microsoft.com/office/b592593e-3fc2-47f2-bec1-bda493811597) | Retorna o número de linhas em uma referência |
| [Função TAXAJURO](https://support.microsoft.com/office/6f5822d8-7ef1-4233-944c-79e8172930f4) | Retorna uma taxa de juros equivalente para o crescimento de um investimento |
| [Função SEC](https://support.microsoft.com/office/ff224717-9c87-4170-9b58-d069ced6d5f7) | Retorna a secante de um ângulo |
| [Função SECH](https://support.microsoft.com/office/e05a789f-5ff7-4d7f-984a-5edb9b09556f) | Retorna a secante hiperbólica de um ângulo |
| [Função SEGUNDO](https://support.microsoft.com/office/740d1cfc-553c-4099-b668-80eaa24e8af1) | Converte um número de série em um segundo |
| [Função SOMASEQÜÊNCIA](https://support.microsoft.com/office/a3ab25b5-1093-4f5b-b084-96c49087f637) | Retorna a soma de uma série polinomial baseada na fórmula |
| [Função PLAN](https://support.microsoft.com/office/44718b6f-8b87-47a1-a9d6-b701c06cff24) | Retorna o número da planilha referenciada |
| [Função PLANS](https://support.microsoft.com/office/770515eb-e1e8-45ce-8066-b557e5e4b80b) | Retorna o número de planilhas em uma referência |
| [Função SINAL](https://support.microsoft.com/office/109c932d-fcdc-4023-91f1-2dd0e916a1d8) | Retorna o sinal de um número |
| [Função SEN](https://support.microsoft.com/office/cf0e3432-8b9e-483c-bc55-a76651c95602) | Retorna o seno do ângulo fornecido |
| [Função SENH](https://support.microsoft.com/office/1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7) | Retorna o seno hiperbólico de um número |
| [Função DISTORÇÃO](https://support.microsoft.com/office/bdf49d86-b1ef-4804-a046-28eaea69c9fa) | Retorna a distorção de uma distribuição |
| [Função DISTORÇÃO.P](https://support.microsoft.com/office/76530a5c-99b9-48a1-8392-26632d542fcb) | Retorna a inclinação de uma distribuição com base em um preenchimento: uma caracterização do grau de assimetria de uma distribuição em torno de sua média |
| [Função DPD](https://support.microsoft.com/office/cdb666e5-c1c6-40a7-806a-e695edc2f1c8) | Retorna a depreciação em linha reta de um ativo durante um período |
| [Função MENOR](https://support.microsoft.com/office/17da8222-7c82-42b2-961b-14c45384df07) | Retorna o menor valor k-ésimo em um conjunto de dados |
| [Função RAIZ](https://support.microsoft.com/office/654975c2-05c4-4831-9a24-2c65e4040fdf) | Retorna uma raiz quadrada positiva |
| [Função RAIZPI](https://support.microsoft.com/office/1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4) | Retorna a raiz quadrada de (número * pi) |
| [Função PADRONIZAR](https://support.microsoft.com/office/81d66554-2d54-40ec-ba83-6437108ee775) | Retorna um valor normalizado |
| [Função DESVPAD.P](https://support.microsoft.com/office/6e917c05-31a0-496f-ade7-4f4e7462f285) | Calcula o desvio padrão com base no preenchimento completo |
| [Função DESVPAD.A](https://support.microsoft.com/office/7d69cf97-0c1f-4acf-be27-f3e83904cc23) | Estima o desvio padrão com base em uma amostra |
| [Função DESVPADA](https://support.microsoft.com/office/5ff38888-7ea5-48de-9a6d-11ed73b29e9d) | Estima o desvio padrão com base em uma amostra, incluindo números, texto e valores lógicos |
| [Função DESVPADPA](https://support.microsoft.com/office/5578d4d6-455a-4308-9991-d405afe2c28c) | Calcula o desvio padrão com base no preenchimento completo, incluindo números, texto e valores lógicos |
| [Função SUBSTITUIR](https://support.microsoft.com/office/6434944e-a904-4336-a9b0-1e58df3bc332) | Substitui um novo texto por um texto antigo em uma cadeia de texto |
| [Função SUBTOTAL](https://support.microsoft.com/office/7b027003-f060-4ade-9040-e478765b9939) | Retorna um subtotal em uma lista ou banco de dados |
| [Função SOMA](https://support.microsoft.com/office/043e1c7d-7726-4e80-8f32-07b23e057f89) | Soma seus argumentos |
| [Função SOMASE](https://support.microsoft.com/office/169b8c99-c05c-4483-a712-1697a653039b) | Adiciona as células especificadas por um determinado critério |
| [Função SOMASES](https://support.microsoft.com/office/c9e748f5-7ea7-455d-9406-611cebce642b) | Adiciona as células de um intervalo que atendam a vários critérios |
| [Função SOMAQUAD](https://support.microsoft.com/office/e3313c02-51cc-4963-aae6-31442d9ec307) | Retorna a soma dos quadrados dos argumentos |
| [Função SDA](https://support.microsoft.com/office/069f8106-b60b-4ca2-98e0-2a0f206bdb27) | Retorna a depreciação dos dígitos da soma dos anos de um ativo para um período especificado |
| [Função T](https://support.microsoft.com/office/fb83aeec-45e7-4924-af95-53e073541228) | Converte os argumentos em texto |
| [Função DIST.T](https://support.microsoft.com/office/4329459f-ae91-48c2-bba8-1ead1c6c21b2) | Retorna os Pontos Percentuais (probabilidade) para a distribuição t de Student |
| [Função DIST.T.BC](https://support.microsoft.com/office/198e9340-e360-4230-bd21-f52f22ff5c28) | Retorna os Pontos Percentuais (probabilidade) para a distribuição t de Student |
| [Função DIST.T.CD](https://support.microsoft.com/office/20a30020-86f9-4b35-af1f-7ef6ae683eda) | Retorna a distribuição t de Student |
| [Função INV.T](https://support.microsoft.com/office/2908272b-4e61-4942-9df9-a25fec9b0e2e) | Retorna o valor t da distribuição t de Student como uma função da probabilidade e dos graus de liberdade |
| [Função INV.T.BC](https://support.microsoft.com/office/ce72ea19-ec6c-4be7-bed2-b9baf2264f17) | Retorna o inverso da distribuição t de Student |
| [Função TAN](https://support.microsoft.com/office/08851a40-179f-4052-b789-d7f699447401) | Retorna a tangente de um número |
| [Função TANH](https://support.microsoft.com/office/017222f0-a0c3-4f69-9787-b3202295dc6c) | Retorna a tangente hiperbólica de um número |
| [Função OTN](https://support.microsoft.com/office/2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c) | Retorna o rendimento de um título equivalente a uma obrigação do Tesouro |
| [Função OTNVALOR](https://support.microsoft.com/office/eacca992-c29d-425a-9eb8-0513fe6035a2) | Retorna o preço por R$ 100,00 do valor nominal de uma obrigação do Tesouro |
| [Função OTNLUCRO](https://support.microsoft.com/office/6d381232-f4b0-4cd5-8e97-45b9c03468ba) | Retorna o rendimento de uma obrigação do Tesouro |
| [Função TEXTO](https://support.microsoft.com/office/20d5ac4d-7b94-49fd-bb38-93d29371225c) | Formata um número e o converte em texto |
| [Função TEMPO](https://support.microsoft.com/office/9a5aff99-8f7d-4611-845e-747d0b8d5457) | Retorna o número de série de uma hora específica |
| [Função VALOR.TEMPO](https://support.microsoft.com/office/0b615c12-33d8-4431-bf3d-f3eb6d186645) | Converte um horário na forma de texto em um número de série |
| [Função HOJE](https://support.microsoft.com/office/5eb3078d-a82c-4736-8930-2f51a028fdd9) | Retorna o número de série da data de hoje |
| [Função ARRUMAR](https://support.microsoft.com/office/410388fa-c5df-49c6-b16c-9e5630b479f9) | Remove espaços do texto |
| [Função MÉDIA.INTERNA](https://support.microsoft.com/office/d90c9878-a119-4746-88fa-63d988f511d3) | Retorna a média do interior de um conjunto de dados |
| [Função VERDADEIRO](https://support.microsoft.com/office/7652c6e3-8987-48d0-97cd-ef223246b3fb) | Retorna o valor lógico `TRUE` |
| [Função TRUNC](https://support.microsoft.com/office/8b86a64c-3127-43db-ba14-aa5ceb292721) | Trunca um número em um inteiro |
| [Função TIPO](https://support.microsoft.com/office/45b4e688-4bc3-48b3-a105-ffa892995899) | Retorna um número indicando o tipo de dados de um valor |
| [Função CARACTUNICODE](https://support.microsoft.com/office/ffeb64f5-f131-44c6-b332-5cd72f0659b8) | Retorna o caractere Unicode referenciado por determinado valor numérico |
| [Função UNICODE](https://support.microsoft.com/office/adb74aaa-a2a5-4dde-aff6-966e4e81f16f) | Retorna o número (ponto de código) que corresponde ao primeiro caractere do texto |
| [Função MAIÚSCULA](https://support.microsoft.com/office/c11f29b3-d1a3-4537-8df6-04d0049963d6) | Converte texto em maiúsculas |
| [Função VALOR](https://support.microsoft.com/office/257d0108-07dc-437d-ae1c-bc2d3953d8c2) | Converte um argumento de texto em um número |
| [Função VAR.P](https://support.microsoft.com/office/73d1285c-108c-4843-ba5d-a51f90656f3a) | Calcula a variação com base no preenchimento inteiro |
| [Função VAR.A](https://support.microsoft.com/office/913633de-136b-449d-813e-65a00b2b990b) | Estima a variação com base em uma amostra |
| [Função VARA](https://support.microsoft.com/office/3de77469-fa3a-47b4-85fd-81758a1e1d07) | Estima a variação com base em uma amostra, incluindo números, texto e valores lógicos |
| [Função VARPA](https://support.microsoft.com/office/59a62635-4e89-4fad-88ac-ce4dc0513b96) | Calcula a variação com base no preenchimento total, incluindo números, texto e valores lógicos |
| [Função BDV](https://support.microsoft.com/office/dde4e207-f3fa-488d-91d2-66d55e861d73) | Retorna a depreciação de um ativo para um período especificado ou parcial usando um método de balanço declinante |
| [Função PROCV](https://support.microsoft.com/office/0bbc8083-26fe-4963-8ab8-93a18ad188a1) | Procura na primeira coluna de uma matriz e se move ao longo da linha para retornar o valor de uma célula |
| [Função DIA.DA.SEMANA](https://support.microsoft.com/office/60e44483-2ed1-439f-8bd0-e404c190949a) | Converte um número de série em um dia da semana |
| [Função NÚMSEMANA](https://support.microsoft.com/office/e5c43a03-b4ab-426c-b411-b18c13c75340) | Converte um número de série em um número que representa onde a semana cai numericamente em um ano |
| [Função DIST.WEIBULL](https://support.microsoft.com/office/4e783c39-9325-49be-bbc9-a83ef82b45db) | Retorna a distribuição de Weibull |
| [Função DIATRABALHO](https://support.microsoft.com/office/f764a5b7-05fc-4494-9486-60d494efbf33) | Retorna o número de série da data antes ou depois de um número específico de dias úteis |
| [Função DIATRABALHO.INTL](https://support.microsoft.com/office/a378391c-9ba7-4678-8a39-39611a9bf81d) | Retorna o número de série da data antes ou depois de um número específico de dias úteis usando parâmetros para indicar quais e quantos dias são de fim de semana |
| [Função XTIR](https://support.microsoft.com/office/de1242ec-6477-445b-b11b-a303ad9adc9d) | Fornece a taxa interna de retorno para um programa de fluxos de caixa que não é necessariamente periódico |
| [Função XVPL](https://support.microsoft.com/office/1b42bbf6-370f-4532-a0eb-d67c16b664b7) | Retorna o valor presente líquido de um programa de fluxos de caixa que não é necessariamente periódico |
| [Função XOR](https://support.microsoft.com/office/1548d4c2-5e47-4f77-9a92-0533bba14f37) | Retorna um OU exclusivo lógico de todos os argumentos |
| [Função ANO](https://support.microsoft.com/office/c64f017a-1354-490d-981f-578e8ec8d3b9) | Converte um número de série em um ano |
| [Função FRAÇÃOANO](https://support.microsoft.com/office/3844141e-c76d-4143-82b6-208454ddc6a8) | Retorna a fração do ano que representa o número de dias entre a data_inicial e a data_final |
| [Função LUCRO](https://support.microsoft.com/office/f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe) | Retorna o lucro de um título que paga juros periódicos |
| [Função LUCRODESC](https://support.microsoft.com/office/a9dbdbae-7dae-46de-b995-615faffaaed7) | Retorna o rendimento anual de um título descontado. Por exemplo, uma obrigação do Tesouro |
| [Função LUCROVENC](https://support.microsoft.com/office/ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f) | Retorna o rendimento anual de um título que paga juros no vencimento |
| [Função TESTE.Z](https://support.microsoft.com/office/d633d5a3-2031-4614-a016-92180ad82bee) | Retorna o valor de probabilidade unicaudal de um teste-z |

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Classe Functions (API JavaScript para Excel)](/javascript/api/excel/excel.functions)
- [Objeto Workbook Functions (API JavaScript para Excel)](/javascript/api/excel/excel.workbook#excel-excel-workbook-functions-member)
