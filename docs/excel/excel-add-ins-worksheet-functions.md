---
title: Chamar funções internas de planilha do Excel usando as APIs JavaScript do Excel
description: ''
ms.date: 01/24/2017
localization_priority: Normal
ms.openlocfilehash: 5ce8ac0c56a7d6a499f601fcc0767a1e76ea14cc
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388609"
---
# <a name="call-built-in-excel-worksheet-functions"></a>Chamar funções internas de planilha do Excel

Este artigo explica como chamar funções internas de planilha do Excel, como `VLOOKUP` e `SUM`, usando as API JavaScript do Excel. Também fornece a lista completa de funções internas de planilha Excel que podem ser chamadas usando as APIs JavaScript do Excel.

> [!NOTE]
> Para saber mais sobre como criar *funções personalizadas* no Excel usando as APIs JavaScript do Excel, confira [Criar funções personalizadas no Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Chamar uma função de planilha

O trecho de código a seguir mostra como chamar uma função de planilha, onde `sampleFunction()` é um espaço reservado que deve ser substituído pelo nome da função a chamar e os parâmetros de entrada que a função exige. A propriedade **valor** do objeto **FunctionResult** que uma função de planilha retorna contém o resultado da função especificada. Como mostra este exemplo, você deve carregar `load` a propriedade **valor** do objeto **FunctionResult** antes de lê-lo. Neste exemplo, o resultado da função está simplesmente sendo gravado no console. 

```js
var functionResult = context.workbook.functions.sampleFunction(); 
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> Confira a seção [Funções de planilha com suporte](#supported-worksheet-functions) deste artigo para obter uma lista das funções que podem ser chamadas usando as APIs JavaScript do Excel.

## <a name="sample-data"></a>Dados de amostra

A imagem a seguir mostra uma tabela em uma planilha do Excel com dados de vendas para vários tipos de ferramentas durante um período de três meses. Cada número da tabela representa o número de unidades vendidas de uma ferramenta específica em um mês específico. Os exemplos a seguir mostram como aplicar funções internas da planilha nesses dados.

![Captura de tela de dados de vendas no Excel para martelo, chave inglesa e Serra nos meses novembro, dezembro e janeiro](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Exemplo 1: função individual

O exemplo a seguir se aplica à função `VLOOKUP` para os dados de exemplo descritos anteriormente a fim de identificar o número de chaves inglesas vendidas em novembro.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var unitSoldInNov = context.workbook.functions.vlookup("Wrench", range, 2, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November = ' + unitSoldInNov.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="example-2-nested-functions"></a>Exemplo 2: funções aninhadas

O exemplo de código a seguir aplica a função `VLOOKUP` nos dados de amostras descritos anteriormente para identificar o número de chaves inglesas vendidas em novembro e em dezembro e, em seguida, aplica a função `SUM` para calcular o total de chaves inglesas vendido durante esses dois meses. 

Como mostra este exemplo, quando uma ou mais chamadas de função são aninhadas dentro de outra chamada de função, você só precisa dar `load` no resultado final caso você queira ler (neste exemplo, `sumOfTwoLookups`). Os resultados intermediários (neste exemplo, o resultado de cada função `VLOOKUP`) serão calculados e usados para calcular o resultado final.

```js
Excel.run(function (context) {
    var range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
    var sumOfTwoLookups = context.workbook.functions.sum(
        context.workbook.functions.vlookup("Wrench", range, 2, false), 
        context.workbook.functions.vlookup("Wrench", range, 3, false)
    );
    sumOfTwoLookups.load('value');

    return context.sync()
        .then(function () {
            console.log(' Number of wrenches sold in November and December = ' + sumOfTwoLookups.value);
        });
}).catch(errorHandlerFunction);
```

## <a name="supported-worksheet-functions"></a>Funções de planilha com suporte

As seguintes funções internas de planilhas do Excel podem ser chamadas usando as APIs JavaScript do Excel. 

| Função | Tipo de retorno | Descrição |
|:---------------|:-------------|:-----------|
| <a href="https://support.office.com/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">Função ABS</a> | FunctionResult | Retorna o valor absoluto de um número |
| <a href="https://support.office.com/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">Função JUROSACUM</a> | FunctionResult | Retorna juros acumulados de um título que paga juros periódicos |
| <a href="https://support.office.com/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">Função JUROSACUMV</a> | FunctionResult | Retorna juros acumulados de um título que paga juros no vencimento |
| <a href="https://support.office.com/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">Função ACOS</a> | FunctionResult | Retorna o arco cosseno de um número |
| <a href="https://support.office.com/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">Função ACOSH</a> | FunctionResult | Retorna o cosseno hiperbólico inverso de um número |
| <a href="https://support.office.com/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">Função ACOT</a> | FunctionResult | Retorna o arco cotangente de um número |
| <a href="https://support.office.com/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">Função ACOTH</a> | FunctionResult | Retorna o arco cotangente hiperbólico de um número |
| <a href="https://support.office.com/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">Função AMORDEGRC</a> | FunctionResult | Retorna a depreciação para cada período contábil usando o coeficiente de depreciação |
| <a href="https://support.office.com/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">Função AMORLINC</a> | FunctionResult | Retorna a depreciação para cada período contábil |
| <a href="https://support.office.com/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">Função E</a> | FunctionResult | Retorna `TRUE` se todos os argumentos forem verdadeiros |
| <a href="https://support.office.com/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">Função ARÁBICO</a> | FunctionResult | Converte um número romano em arábico, como um número |
| <a href="https://support.office.com/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">Função ÁREAS</a> | FunctionResult | Retorna o número de áreas em uma referência |
| <a href="https://support.office.com/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">Função ASC</a> | FunctionResult | Altera letras do inglês ou katakana de largura total (bytes duplos) dentro de uma cadeia de caracteres para caracteres de meia largura (byte único) |
| <a href="https://support.office.com/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">Função ASEN</a> | FunctionResult | Retorna o arco seno de um número |
| <a href="https://support.office.com/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c" target="_blank">Função ASENH</a> | FunctionResult | Retorna o seno hiperbólico inverso de um número |
| <a href="https://support.office.com/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">Função ATAN</a> | FunctionResult | Retorna o arco tangente de um número |
| <a href="https://support.office.com/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">Função ATAN2</a> | FunctionResult | Retorna o arco tangente das coordenadas x e y especificadas |
| <a href="https://support.office.com/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">Função ATANH</a> | FunctionResult | Retorna a tangente hiperbólica inversa de um número |
| <a href="https://support.office.com/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">Função DESV.MÉDIO</a> | FunctionResult | Retorna a média dos desvios absolutos dos pontos de dados a partir de sua média |
| <a href="https://support.office.com/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">Função MÉDIA</a> | FunctionResult | Retorna a média dos argumentos |
| <a href="https://support.office.com/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">Função MÉDIAA</a> | FunctionResult | Retorna a média dos argumentos, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">Função MÉDIASE</a> | FunctionResult | Retorna a média (média aritmética) de todas as células em um intervalo que atendem a um determinado critério |
| <a href="https://support.office.com/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">Função MÉDIASES</a> | FunctionResult | Retorna a média (média aritmética) de todas as células que satisfazem vários critérios |
| <a href="https://support.office.com/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">Função BAHTTEXT</a> | FunctionResult | Converte um número em texto, usando o formato de moeda ß (baht) |
| <a href="https://support.office.com/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">Função BASE</a> | FunctionResult | Converte um número em uma representação de texto com a determinada base |
| <a href="https://support.office.com/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">Função BESSELI</a> | FunctionResult | Retorna a função de Bessel In(x) modificada |
| <a href="https://support.office.com/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">Função BESSELJ</a> | FunctionResult | Retorna a função de Bessel Jn(x) |
| <a href="https://support.office.com/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">Função BESSELK</a> | FunctionResult | Retorna a função de Bessel Kn(x) modificada |
| <a href="https://support.office.com/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">Função BESSELY</a> | FunctionResult | Retorna a função de Bessel Yn(x) |
| <a href="https://support.office.com/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">Função DIST.BETA</a> | FunctionResult | Retorna a função de distribuição cumulativa beta |
| <a href="https://support.office.com/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">Função INV.BETA</a> | FunctionResult | Retorna o inverso da função de distribuição cumulativa para uma distribuição beta especificada |
| <a href="https://support.office.com/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">Função BIN2DEC</a> | FunctionResult | Converte um número binário em decimal |
| <a href="https://support.office.com/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">Função BIN2HEX</a> | FunctionResult | Converte um número binário em hexadecimal |
| <a href="https://support.office.com/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">Função BIN2OCT</a> | FunctionResult | Converte um número binário em octal |
| <a href="https://support.office.com/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">Função DISTR.BINOM</a> | FunctionResult | Retorna a probabilidade de distribuição binomial do termo individual |
| <a href="https://support.office.com/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595" target="_blank">Função INTERV.DISTR.BINOM</a> | FunctionResult | Retorna a probabilidade de um resultado de teste usando uma distribuição binomial |
| <a href="https://support.office.com/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">Função INV.BINOM</a> | FunctionResult | Retorna o menor valor para o qual a distribuição binomial cumulativa é maior ou igual ao valor padrão |
| <a href="https://support.office.com/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">Função BITAND</a> | FunctionResult | Retorna um bit "E" de dois números |
| <a href="https://support.office.com/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">Função DESLOCESQBIT</a> | FunctionResult | Retorna um valor numérico deslocado à esquerda por quantidade_deslocamento bits |
| <a href="https://support.office.com/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">Função BITOR</a> | FunctionResult | Retorna um bit "OU" de dois números |
| <a href="https://support.office.com/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">Função DESLOCDIRBIT</a> | FunctionResult | Retorna um valor numérico deslocado à direita por quantidade_deslocamento bits |
| <a href="https://support.office.com/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">Função BITXOR</a> | FunctionResult | Retorna um bit 'Exclusivo Ou' de dois números |
| <a href="https://support.office.com/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">TETO. MATEMÁTICA, funções ECMA_CEILING</a> | FunctionResult | Arredonda um número para cima, para o inteiro mais próximo ou para o múltiplo mais próximo significativo |
| <a href="https://support.office.com/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">Função TETO.PRECISO</a> | FunctionResult | Arredonda um número para o inteiro mais próximo ou para o múltiplo mais próximo significativo. Independentemente do sinal do número, ele é arredondado para cima. |
| <a href="https://support.office.com/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">Função CARACT</a> | FunctionResult | Retorna o caractere especificado pelo número de código |
| <a href="https://support.office.com/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">Função DIST.QUIQUA</a> | FunctionResult | Retorna a função de densidade da probabilidade beta cumulativa |
| <a href="https://support.office.com/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">Função DIST.QUIQUA.CD</a> | FunctionResult | Retorna a probabilidade unicaudal da distribuição qui-quadrada |
| <a href="https://support.office.com/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f" target="_blank">Função INV.QUIQUA</a> | FunctionResult | Retorna a função de densidade da probabilidade beta cumulativa |
| <a href="https://support.office.com/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">Função INV.QUIQUA.CD</a> | FunctionResult | Retorna o inverso da probabilidade unicaudal da distribuição qui-quadrada |
| <a href="https://support.office.com/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">Função ESCOLHER</a> | FunctionResult | Seleciona um valor em uma lista de valores |
| <a href="https://support.office.com/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">Função TIRAR</a> | FunctionResult | Remove do texto todos os caracteres não imprimíveis |
| <a href="https://support.office.com/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">Função CÓDIGO</a> | FunctionResult | Retorna um código numérico para o primeiro caractere de uma cadeia de texto |
| <a href="https://support.office.com/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">Função COLS</a> | FunctionResult | Retorna o número de colunas em uma referência |
| <a href="https://support.office.com/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">Função COMBIN</a> | FunctionResult | Retorna o número de combinações de um determinado número de objetos |
| <a href="https://support.office.com/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">Função COMBINA</a> | FunctionResult | Retorna o número de combinações com repetições de um determinado número de itens |
| <a href="https://support.office.com/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">Função COMPLEXO</a> | FunctionResult | Converte coeficientes reais e imaginários em um número complexo |
| <a href="https://support.office.com/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">Função CONCATENAR</a> | FunctionResult | Agrupa vários itens de texto em um item de texto |
| <a href="https://support.office.com/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">Função INT.CONFIANÇA.NORM</a> | FunctionResult | Retorna o intervalo de confiança para um meio de preenchimento |
| <a href="https://support.office.com/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">Função INT.CONFIANÇA.T</a> | FunctionResult | Retorna o intervalo de confiança para um meio de preenchimento, usando a distribuição t de Student |
| <a href="https://support.office.com/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">Função CONVERTER</a> | FunctionResult | Converte um número de um sistema de medidas para outro |
| <a href="https://support.office.com/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">Função COS</a> | FunctionResult | Retorna o cosseno de um número |
| <a href="https://support.office.com/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">Função COSH</a> | FunctionResult | Retorna o cosseno hiperbólico de um número |
| <a href="https://support.office.com/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">Função COT</a> | FunctionResult | Retorna a cotangente de um ângulo |
| <a href="https://support.office.com/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">Função COTH</a> | FunctionResult | Retorna a cotangente hiperbólica de um número |
| <a href="https://support.office.com/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">Função CONT.NÚM</a> | FunctionResult | Calcula quantos números há na lista de argumentos |
| <a href="https://support.office.com/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">Função CONT.VALORES</a> | FunctionResult | Calcula quantos valores há na lista de argumentos |
| <a href="https://support.office.com/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">Função CONTAR.VAZIO</a> | FunctionResult | Conta o número de células vazias no intervalo especificado |
| <a href="https://support.office.com/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">Função CONT.SE</a> | FunctionResult | Conta o número de células em um intervalo que atendem aos critérios fornecidos |
| <a href="https://support.office.com/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">Função CONT.SES</a> | FunctionResult | Conta o número de células dentro de um intervalo que atende a múltiplos critérios |
| <a href="https://support.office.com/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">Função CUPDIASINLIQ</a> | FunctionResult | Retorna o número de dias do início do período de cupom até a data de liquidação |
| <a href="https://support.office.com/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">Função CUPDIAS</a> | FunctionResult | Retorna o número de dias no período de cupom que contém a data de liquidação |
| <a href="https://support.office.com/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">Função CUPDIASPRÓX</a> | FunctionResult | Retorna o número de dias da data de liquidação até a data do próximo cupom |
| <a href="https://support.office.com/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">Função CUPDATAPRÓX</a> | FunctionResult | Retorna a próxima data de cupom após a data de quitação |
| <a href="https://support.office.com/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">Função CUPNÚM</a> | FunctionResult | Retorna o número de cupons pagáveis entre as datas de quitação e vencimento |
| <a href="https://support.office.com/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">Função CUPDATAANT</a> | FunctionResult | Retorna a data de cupom anterior à data de quitação |
| <a href="https://support.office.com/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1" target="_blank">Função COSEC</a> | FunctionResult | Retorna a cossecante de um ângulo |
| <a href="https://support.office.com/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">Função COSECH</a> | FunctionResult | Retorna a cossecante hiperbólica de um ângulo |
| <a href="https://support.office.com/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">Função PGTOJURACUM</a> | FunctionResult | Retorna os juros acumulados pagos entre dois períodos |
| <a href="https://support.office.com/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">Função PGTOCAPACUM</a> | FunctionResult | Retorna o capital acumulado pago sobre um empréstimo entre dois períodos |
| <a href="https://support.office.com/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">Função DATA</a> | FunctionResult | Retorna o número de série de uma data específica |
| <a href="https://support.office.com/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">Função DATA.VALOR</a> | FunctionResult | Converte uma data na forma de texto em um número de série |
| <a href="https://support.office.com/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">Função BDMÉDIA</a> | FunctionResult | Retorna a média das entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">Função DIA</a> | FunctionResult | Converte um número de série em um dia do mês |
| <a href="https://support.office.com/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df" target="_blank">Função DIAS</a> | FunctionResult | Retorna o número de dias entre duas datas |
| <a href="https://support.office.com/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">Função DIAS360</a> | FunctionResult | Calcula o número de dias entre duas datas com base em um ano de 360 dias |
| <a href="https://support.office.com/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">Função BD</a> | FunctionResult | Retorna a depreciação de um ativo para um período especificado, usando o método de balanço de declínio fixo |
| <a href="https://support.office.com/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">Função DBCS</a> | FunctionResult | Altera letras do inglês ou katakana de meia largura (byte único) dentro de uma cadeia de caracteres para caracteres de largura total (bytes duplos) |
| <a href="https://support.office.com/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">Função BDCONTAR</a> | FunctionResult | Conta as células que contêm números em um banco de dados |
| <a href="https://support.office.com/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">Função BDCONTARA</a> | FunctionResult | Conta células não vazias em um banco de dados |
| <a href="https://support.office.com/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">Função BDD</a> | FunctionResult | Retorna a depreciação de um ativo com relação a um período especificado usando o método de saldos decrescentes duplos ou qualquer outro método especificado por você |
| <a href="https://support.office.com/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">Função DEC2BIN</a> | FunctionResult | Converte um número decimal em binário |
| <a href="https://support.office.com/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">Função DEC2HEX</a> | FunctionResult | Converte um número decimal em hexadecimal |
| <a href="https://support.office.com/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">Função DEC2OCT</a> | FunctionResult | Converte um número decimal em octal |
| <a href="https://support.office.com/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e" target="_blank">Função DECIMAL</a> | FunctionResult | Converte em um número decimal a representação de texto de um número em determinada base |
| <a href="https://support.office.com/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">Função GRAUS</a> | FunctionResult | Converte radianos em graus |
| <a href="https://support.office.com/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">Função DELTA</a> | FunctionResult | Testa se dois valores são iguais |
| <a href="https://support.office.com/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">Função DESVQ</a> | FunctionResult | Retorna a soma dos quadrados dos desvios |
| <a href="https://support.office.com/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">Função BDEXTRAIR</a> | FunctionResult | Extrai de um banco de dados um único registro que corresponde aos critérios especificados |
| <a href="https://support.office.com/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">Função DESC</a> | FunctionResult | Retorna a taxa de desconto de um título |
| <a href="https://support.office.com/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">Função BDMÁX</a> | FunctionResult | Retorna o valor máximo de entradas selecionadas de banco de dados |
| <a href="https://support.office.com/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">Função BDMÍN</a> | FunctionResult | Retorna o valor mínimo de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">DÓLAR, funções USDOLLAR</a> | FunctionResult | Converte um número em texto, usando o formato de moeda $ (cifrão) |
| <a href="https://support.office.com/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427" target="_blank">Função MOEDADEC</a> | FunctionResult | Converte um preço em moeda expresso como uma fração em um preço em moeda expresso como um número decimal |
| <a href="https://support.office.com/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">Função MOEDAFRA</a> | FunctionResult | Converte um preço em moeda expresso como um número decimal em um preço em moeda expresso como uma fração |
| <a href="https://support.office.com/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">Função BDMULTIPL</a> | FunctionResult | Multiplica os valores em um campo específico de registros que correspondem ao critério em um banco de dados |
| <a href="https://support.office.com/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">Função BDEST</a> | FunctionResult | Estima o desvio padrão com base em uma amostra de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">Função BDDESVPA</a> | FunctionResult | Calcula o desvio padrão com base no preenchimento completo de entradas selecionadas de banco de dados |
| <a href="https://support.office.com/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">Função BDSOMA</a> | FunctionResult | Soma os números na coluna de campos de registros do banco de dados que correspondem ao critério |
| <a href="https://support.office.com/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">Função DURAÇÃO</a> | FunctionResult | Retorna a duração anual de um título com pagamentos de juros periódicos |
| <a href="https://support.office.com/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">Função BDVAREST</a> | FunctionResult | Estima a variação com base em uma amostra de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">Função BDVARP</a> | FunctionResult | Calcula a variação com base no preenchimento completo de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">Função DATAM</a> | FunctionResult | Retorna o número de série da data que é o número indicado de meses antes ou depois da data inicial |
| <a href="https://support.office.com/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">Função EFETIVA</a> | FunctionResult | Retorna a taxa de juros anual efetiva |
| <a href="https://support.office.com/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">Função FIMMÊS</a> | FunctionResult | Retorna o número de série do último dia do mês antes ou depois de um número especificado de meses |
| <a href="https://support.office.com/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">Função FUNERRO</a> | FunctionResult | Retorna a função de erro |
| <a href="https://support.office.com/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">Função FUNERRO.PRECISO</a> | FunctionResult | Retorna a função de erro |
| <a href="https://support.office.com/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">Função FUNERROCOMPL</a> | FunctionResult | Retorna a função de erro complementar |
| <a href="https://support.office.com/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">Função FUNERROCOMPL.PRECISO</a> | FunctionResult | Retorna a função FUNERRO complementar integrada entre x e infinito |
| <a href="https://support.office.com/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">Função TIPO.ERRO</a> | FunctionResult | Retorna um número correspondente a um tipo de erro |
| <a href="https://support.office.com/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">Função PAR</a> | FunctionResult | Arredonda um número para cima até o inteiro par mais próximo |
| <a href="https://support.office.com/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926" target="_blank">Função EXATO</a> | FunctionResult | Verifica se dois valores de texto são idênticos |
| <a href="https://support.office.com/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">Função EXP</a> | FunctionResult | Retorna e elevado à potência de um número especificado |
| <a href="https://support.office.com/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">Função DISTR.EXPON</a> | FunctionResult | Retorna a distribuição exponencial |
| <a href="https://support.office.com/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">Função DIST.F</a> | FunctionResult | Retorna a distribuição de probabilidade F |
| <a href="https://support.office.com/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">Função DIST.F.CD</a> | FunctionResult | Retorna a distribuição de probabilidade F |
| <a href="https://support.office.com/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">Função INV.F</a> | FunctionResult | Retorna o inverso da distribuição de probabilidade F |
| <a href="https://support.office.com/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">Função INV.F.CD</a> | FunctionResult | Retorna o inverso da distribuição de probabilidade F |
| <a href="https://support.office.com/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">Função FATORIAL</a> | FunctionResult | Retorna o fatorial de um número |
| <a href="https://support.office.com/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">Função FATDUPLO</a> | FunctionResult | Retorna o fatorial duplo de um número |
| <a href="https://support.office.com/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">Função FALSO</a> | FunctionResult | Retorna o valor lógico `FALSE` |
| <a href="https://support.office.com/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">Funções PROCURAR, PROCURARB</a> | FunctionResult | Procura um valor de texto dentro de outro (diferencia maiúsculas de minúsculas) |
| <a href="https://support.office.com/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">Função FISHER</a> | FunctionResult | Retorna a transformação Fisher |
| <a href="https://support.office.com/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">Função FISHERINV</a> | FunctionResult | Retorna o inverso da transformação Fisher |
| <a href="https://support.office.com/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">Função FIXO</a> | FunctionResult | Formata um número como texto com um número fixo de decimais |
| <a href="https://support.office.com/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">Função de ARREDMULTB.MAT</a> | FunctionResult | Arredonda um número para baixo para o inteiro mais próximo ou para o múltiplo mais próximo de significância |
| <a href="https://support.office.com/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">Função ARREDMULTB.PRECISO</a> | FunctionResult | Arredonda um número para baixo para o inteiro mais próximo ou para o múltiplo mais próximo de significância. Independentemente do sinal do número, ele é arredondado para baixo. |
| <a href="https://support.office.com/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">Função VF</a> | FunctionResult | Retorna o valor futuro de um investimento |
| <a href="https://support.office.com/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">Função VFPLANO</a> | FunctionResult | Retorna o valor futuro de um capital inicial após a aplicação de uma série de taxas de juros compostas |
| <a href="https://support.office.com/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">Função GAMA</a> | FunctionResult | Retorna o valor da função GAMA |
| <a href="https://support.office.com/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">Função DIST.GAMA</a> | FunctionResult | Retorna a distribuição gama |
| <a href="https://support.office.com/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">Função INV.GAMA</a> | FunctionResult | Retorna o inverso da distribuição cumulativa gama |
| <a href="https://support.office.com/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">Função LNGAMA</a> | FunctionResult | Retorna o logaritmo natural da função gama, G(x) |
| <a href="https://support.office.com/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">Função LNGAMA.PRECISO</a> | FunctionResult | Retorna o logaritmo natural da função gama, G(x) |
| <a href="https://support.office.com/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">Função GAUSS</a> | FunctionResult | Retorna menos 0,5 que a distribuição cumulativa normal padrão |
| <a href="https://support.office.com/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">Função MDC</a> | FunctionResult | Retorna o máximo divisor comum |
| <a href="https://support.office.com/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">Função MÉDIA.GEOMÉTRICA</a> | FunctionResult | Retorna a média geométrica |
| <a href="https://support.office.com/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">Função DEGRAU</a> | FunctionResult | Testa se um número é maior do que um valor limite |
| <a href="https://support.office.com/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">Função MÉDIA.HARMÔNICA</a> | FunctionResult | Retorna a média harmônica |
| <a href="https://support.office.com/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1" target="_blank">Função HEX2BIN</a> | FunctionResult | Converte um número hexadecimal em binário |
| <a href="https://support.office.com/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">Função HEX2DEC</a> | FunctionResult | Converte um número hexadecimal em decimal |
| <a href="https://support.office.com/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">Função HEX2OCT</a> | FunctionResult | Converte um número hexadecimal em octal |
| <a href="https://support.office.com/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">Função PROCH</a> | FunctionResult | Procura na linha superior de uma matriz e retorna o valor da célula especificada |
| <a href="https://support.office.com/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">Função HORA</a> | FunctionResult | Converte um número de série em um hora |
| <a href="https://support.office.com/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">Função HIPERLINK</a> | FunctionResult | Cria um atalho ou salto que abre um documento armazenado em um servidor de rede, uma intranet ou Internet |
| <a href="https://support.office.com/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">Função DIST.HIPERGEOM.N</a> | FunctionResult | Retorna a distribuição hipergeométrica |
| <a href="https://support.office.com/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">Função SE</a> | FunctionResult | Especifica um teste lógico a ser executado |
| <a href="https://support.office.com/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">Função IMABS</a> | FunctionResult | Retorna o valor absoluto (módulo) de um número complexo |
| <a href="https://support.office.com/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">Função IMAGINÁRIO</a> | FunctionResult | Retorna o coeficiente imaginário de um número complexo |
| <a href="https://support.office.com/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">Função IMARG</a> | FunctionResult | Retorna o argumento teta, um ângulo expresso em radianos |
| <a href="https://support.office.com/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">Função IMCONJ</a> | FunctionResult | Retorna o conjugado complexo de um número complexo |
| <a href="https://support.office.com/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">Função IMCOS</a> | FunctionResult | Retorna o cosseno de um número complexo |
| <a href="https://support.office.com/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">Função IMCOSH</a> | FunctionResult | Retorna o cosseno hiperbólico de um número complexo |
| <a href="https://support.office.com/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">Função IMCOT</a> | FunctionResult | Retorna a cotangente de um número complexo |
| <a href="https://support.office.com/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">Função IMCOSEC</a> | FunctionResult | Retorna a cossecante de um número complexo |
| <a href="https://support.office.com/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">Função IMCOSECH</a> | FunctionResult | Retorna a cossecante hiperbólica de um número complexo |
| <a href="https://support.office.com/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">Função IMDIV</a> | FunctionResult | Retorna o quociente de dois números complexos |
| <a href="https://support.office.com/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">Função IMEXP</a> | FunctionResult | Retorna o exponencial de um número complexo |
| <a href="https://support.office.com/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">Função IMLN</a> | FunctionResult | Retorna o logaritmo natural de um número complexo |
| <a href="https://support.office.com/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">Função IMLOG10</a> | FunctionResult | Retorna o logaritmo de base 10 de um número complexo |
| <a href="https://support.office.com/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">Função IMLOG2</a> | FunctionResult | Retorna o logaritmo de base 2 de um número complexo |
| <a href="https://support.office.com/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">Função IMPOT</a> | FunctionResult | Retorna um número complexo elevado a uma potência inteira |
| <a href="https://support.office.com/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">Função IMPROD</a> | FunctionResult | Retorna o produto de 2 a 255 números complexos |
| <a href="https://support.office.com/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">Função IMREAL</a> | FunctionResult | Retorna o coeficiente real de um número complexo |
| <a href="https://support.office.com/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">Função IMSEC</a> | FunctionResult | Retorna a secante de um número complexo |
| <a href="https://support.office.com/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b" target="_blank">Função IMSECH</a> | FunctionResult | Retorna a secante hiperbólica de um número complexo |
| <a href="https://support.office.com/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">Função IMSENO</a> | FunctionResult | Retorna o seno de um número complexo |
| <a href="https://support.office.com/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">Função IMSENH</a> | FunctionResult | Retorna o seno hiperbólico de um número complexo |
| <a href="https://support.office.com/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">Função IMSQRT</a> | FunctionResult | Retorna a raiz quadrada de um número complexo |
| <a href="https://support.office.com/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">Função IMSUBTR</a> | FunctionResult | Retorna a diferença entre dois números complexos |
| <a href="https://support.office.com/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">Função IMSOMA</a> | FunctionResult | Retorna a soma de números complexos |
| <a href="https://support.office.com/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">Função IMTAN</a> | FunctionResult | Retorna a tangente de um número complexo |
| <a href="https://support.office.com/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">Função INT</a> | FunctionResult | Arredonda um número para baixo até o número inteiro mais próximo |
| <a href="https://support.office.com/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">Função TAXAJUROS</a> | FunctionResult | Retorna a taxa de juros de um título totalmente investido |
| <a href="https://support.office.com/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">Função IPGTO</a> | FunctionResult | Retorna o pagamento de juros para um investimento em um determinado período |
| <a href="https://support.office.com/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">Função TIR</a> | FunctionResult | Retorna a taxa interna de retorno de uma série de fluxos de caixa |
| <a href="https://support.office.com/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉERRO</a> | FunctionResult | Retorna `TRUE` se o valor for qualquer valor de erro, exceto # n/d |
| <a href="https://support.office.com/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉERROS</a> | FunctionResult | Retorna `TRUE` se o valor for qualquer valor de erro |
| <a href="https://support.office.com/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">Função ÉPAR</a> | FunctionResult | Retorna `TRUE` se o número for par |
| <a href="https://support.office.com/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">Função ÉFÓRMULA</a> | FunctionResult | Retorna `TRUE` quando há uma referência a uma célula que contém uma fórmula |
| <a href="https://support.office.com/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉLÓGICO</a> | FunctionResult | Retorna `TRUE` se o valor for um valor lógico |
| <a href="https://support.office.com/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função É.NÃO.DISP</a> | FunctionResult | Retorna `TRUE` se o valor é o valor de erro # n / |
| <a href="https://support.office.com/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função É.NÃO.TEXTO</a> | FunctionResult | Retorna `TRUE` se o valor não for texto |
| <a href="https://support.office.com/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉNÚM</a> | FunctionResult | Retorna `TRUE` se o valor for um número |
| <a href="https://support.office.com/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">Função ISO.TETO</a> | FunctionResult | Retorna um número para o inteiro mais próximo ou para o múltiplo mais próximo de significância |
| <a href="https://support.office.com/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉIMPAR</a> | FunctionResult | Retorna `TRUE` se o número for ímpar |
| <a href="https://support.office.com/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">Função NÚMSEMANAISO</a> | FunctionResult | Retorna o número do número da semana ISO do ano referente a determinada data |
| <a href="https://support.office.com/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">Função ÉPGTO</a> | FunctionResult | Calcula os juros pagos durante um período específico de um investimento |
| <a href="https://support.office.com/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉREF</a> | FunctionResult | Retorna `TRUE` se o valor for uma referência |
| <a href="https://support.office.com/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Função ÉTEXTO</a> | FunctionResult | Retorna `TRUE` se o valor for texto |
| <a href="https://support.office.com/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">Função CURT</a> | FunctionResult | Retorna a curtose de um conjunto de dados |
| <a href="https://support.office.com/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">Função MAIOR</a> | FunctionResult | Retorna o maior valor k-ésimo em um conjunto de dados |
| <a href="https://support.office.com/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">Função MMC</a> | FunctionResult | Retorna o mínimo múltiplo comum |
| <a href="https://support.office.com/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">Funções ESQUERDA, ESQUERDAB</a> | FunctionResult | Retorna os caracteres mais à esquerda de um valor de texto |
| <a href="https://support.office.com/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">Funções NÚM.CARACT, NÚM.CARACTB</a> | FunctionResult | Retorna o número de caracteres em uma cadeia de texto |
| <a href="https://support.office.com/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">Função LN</a> | FunctionResult | Retorna o logaritmo natural de um número |
| <a href="https://support.office.com/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">Função LOG</a> | FunctionResult | Retorna o logaritmo de um número de uma base especificada |
| <a href="https://support.office.com/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">Função LOG10</a> | FunctionResult | Retorna o logaritmo de base 10 de um número |
| <a href="https://support.office.com/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">Função DIST.LOGNORMAL.N</a> | FunctionResult | Retorna a distribuição lognormal cumulativa |
| <a href="https://support.office.com/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">Função INV.LOGNORMAL</a> | FunctionResult | Retorna o inverso da distribuição cumulativa lognormal |
| <a href="https://support.office.com/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb" target="_blank">Função PROC</a> | FunctionResult | Procura valores em um vetor ou matriz |
| <a href="https://support.office.com/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">Função MINÚSCULA</a> | FunctionResult | Converte texto em minúsculas |
| <a href="https://support.office.com/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">Função CORRESP</a> | FunctionResult | Procura valores em uma referência ou matriz |
| <a href="https://support.office.com/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">Função MÁXIMO</a> | FunctionResult | Retorna o valor máximo em uma lista de argumentos |
| <a href="https://support.office.com/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">Função MÁXIMOA</a> | FunctionResult | Retorna o maior valor em uma lista de argumentos, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">Função MDURAÇÃO</a> | FunctionResult | Retorna a duração modificada Macauley de um título com um valor de paridade equivalente a R$ 100 |
| <a href="https://support.office.com/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">Função MED</a> | FunctionResult | Retorna a mediana dos números indicados |
| <a href="https://support.office.com/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">Funções EXT.TEXTO, EXT.TEXTOB</a> | FunctionResult | Retorna um número específico de caracteres de uma cadeia de texto começando na posição especificada |
| <a href="https://support.office.com/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">Função MÍNIMO</a> | FunctionResult | Retorna o valor mínimo em uma lista de argumentos |
| <a href="https://support.office.com/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">Função MÍNIMOA</a> | FunctionResult | Retorna o menor valor em uma lista de argumentos, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">Função MINUTO</a> | FunctionResult | Converte um número de série em um minuto |
| <a href="https://support.office.com/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">Função MTIR</a> | FunctionResult | Calcula a taxa interna de retorno em que fluxos de caixa positivos e negativos são financiados com diferentes taxas |
| <a href="https://support.office.com/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">Função MOD</a> | FunctionResult | Retorna o resto da divisão |
| <a href="https://support.office.com/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">Função MÊS</a> | FunctionResult | Converte um número de série em um mês |
| <a href="https://support.office.com/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">Função MARRED</a> | FunctionResult | Retorna um número arredondado ao múltiplo desejado |
| <a href="https://support.office.com/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">Função MULTINOMIAL</a> | FunctionResult | Retorna o multinômio de um conjunto de números |
| <a href="https://support.office.com/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">Função N</a> | FunctionResult | Retorna um valor convertido em um número |
| <a href="https://support.office.com/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">Função NÃO.DISP</a> | FunctionResult | Retorna o valor de erro # n/d |
| <a href="https://support.office.com/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">Função DIST.BIN.NEG.N</a> | FunctionResult | Retorna a distribuição binomial negativa |
| <a href="https://support.office.com/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">Função DIATRABALHOTOTAL</a> | FunctionResult | Retorna o número de dias úteis inteiros entre duas datas |
| <a href="https://support.office.com/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">Função DIATRABALHOTOTAL.INTL</a> | FunctionResult | Retorna o número de dias de trabalho totais entre duas datas usando parâmetros para indicar quais e quantos dias caem em finais de semana |
| <a href="https://support.office.com/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">Função NOMINAL</a> | FunctionResult | Retorna a taxa de juros anual nominal |
| <a href="https://support.office.com/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">Função DIST.NORM.N</a> | FunctionResult | Retorna a distribuição cumulativa normal |
| <a href="https://support.office.com/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">Função INV.NORM.N</a> | FunctionResult | Retorna o inverso da distribuição cumulativa normal |
| <a href="https://support.office.com/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">Função DIST.NORMP.N</a> | FunctionResult | Retorna a distribuição cumulativa normal padrão |
| <a href="https://support.office.com/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">Função INV.NORMP.N</a> | FunctionResult | Retorna o inverso da distribuição cumulativa normal padrão |
| <a href="https://support.office.com/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">Função NÃO</a> | FunctionResult | Inverte o valor lógico do argumento |
| <a href="https://support.office.com/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">Função AGORA</a> | FunctionResult | Retorna o número de série sequencial da data e hora atuais |
| <a href="https://support.office.com/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">Função NPER</a> | FunctionResult | Retorna o número de períodos de um investimento |
| <a href="https://support.office.com/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">Função VPL</a> | FunctionResult | Retorna o valor líquido atual de um investimento com base em uma série de fluxos de caixa periódicos e em uma taxa de desconto |
| <a href="https://support.office.com/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">Função VALORNUMÉRICO</a> | FunctionResult | Converte texto em número de maneira independente de localidade |
| <a href="https://support.office.com/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">Função OCT2BIN</a> | FunctionResult | Converte um número octal em binário |
| <a href="https://support.office.com/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">Função OCT2DEC</a> | FunctionResult | Converte um número octal em decimal |
| <a href="https://support.office.com/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f" target="_blank">Função OCT2HEX</a> | FunctionResult | Converte um número octal em hexadecimal |
| <a href="https://support.office.com/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">Função ÍMPAR</a> | FunctionResult | Arredonda um número para cima até o inteiro ímpar mais próximo |
| <a href="https://support.office.com/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">Função PREÇOPRIMINC</a> | FunctionResult | Retorna o preço por R$ 100 do valor nominal de um título com um período inicial incompleto |
| <a href="https://support.office.com/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">Função LUCROPRIMINC</a> | FunctionResult | Retorna o rendimento de um título com um período inicial incompleto |
| <a href="https://support.office.com/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">Função PREÇOÚLTINC</a> | FunctionResult | Retorna o preço por R$ 100 do valor nominal de um título com um período final incompleto |
| <a href="https://support.office.com/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">Função LUCROÚLTINC</a> | FunctionResult | Retorna o rendimento de um título com um período final incompleto |
| <a href="https://support.office.com/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">Função OU</a> | FunctionResult | Retorna `TRUE` se algum argumento não for true |
| <a href="https://support.office.com/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf" target="_blank">Função DURAÇÃOP</a> | FunctionResult | Retorna o número de períodos necessários para que um investimento atinja um valor específico |
| <a href="https://support.office.com/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">Função PERCENTIL.EXC</a> | FunctionResult | Retorna o k-ésimo percentil de valores em um intervalo, onde k está no intervalo 0..1, exclusive |
| <a href="https://support.office.com/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">Função PERCENTIL.INC</a> | FunctionResult | Retorna o k-ésimo percentil de valores em um intervalo |
| <a href="https://support.office.com/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">Função ORDEM.PORCENTUAL.EXC</a> | FunctionResult | Retorna a posição de um valor em um conjunto de dados como uma porcentagem (0..1, exclusivo) do conjunto de dados |
| <a href="https://support.office.com/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">Função ORDEM.PORCENTUAL.INC</a> | FunctionResult | Retorna a posição de porcentagem de um valor em um conjunto de dados |
| <a href="https://support.office.com/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">Função PERMUT</a> | FunctionResult | Retorna o número de permutações de um determinado número de objetos |
| <a href="https://support.office.com/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">Função PERMUTAS</a> | FunctionResult | Retorna o número de permutações referentes a determinado número de objetos (com repetições) que podem ser selecionadas do total de objetos |
| <a href="https://support.office.com/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">Função PHI</a> | FunctionResult | Retorna o valor da função de densidade referente a uma distribuição normal padrão |
| <a href="https://support.office.com/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">Função PI</a> | FunctionResult | Retorna o valor de pi |
| <a href="https://support.office.com/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441" target="_blank">Função PGTO</a> | FunctionResult | Retorna o pagamento periódico de uma anuidade |
| <a href="https://support.office.com/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">Função DIST.POISSON</a> | FunctionResult | Retorna a distribuição de Poisson |
| <a href="https://support.office.com/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">Função POTÊNCIA</a> | FunctionResult | Retorna o resultado de um número elevado a uma potência |
| <a href="https://support.office.com/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">Função PPGTO</a> | FunctionResult | Retorna o pagamento de capital para determinado período de investimento |
| <a href="https://support.office.com/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">Função PREÇO</a> | FunctionResult | Retorna o preço pelo valor nominal R$100 de um título que paga juros periódicos |
| <a href="https://support.office.com/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">Função PREÇODESC</a> | FunctionResult | Retorna o preço por valor nominal de R$ 100,00 de um título descontado |
| <a href="https://support.office.com/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">Função PREÇOVENC</a> | FunctionResult | Retorna o preço pelo valor nominal R$100 de um título que paga juros no vencimento |
| <a href="https://support.office.com/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">Função MULT</a> | FunctionResult | Multiplica os argumentos |
| <a href="https://support.office.com/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">Função PRI.MAIÚSCULA</a> | FunctionResult | Coloca a primeira letra de cada palavra em maiúscula em um valor de texto |
| <a href="https://support.office.com/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">Função VP</a> | FunctionResult | Retorna o valor presente de um investimento |
| <a href="https://support.office.com/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">Função QUARTIL.EXC</a> | FunctionResult | Retorna o quartil do conjunto de dados, com base em valores de percentil de 0..1, exclusive. |
| <a href="https://support.office.com/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">Função QUARTIL.INC</a> | FunctionResult | Retorna o quartil do conjunto de dados |
| <a href="https://support.office.com/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">Função QUOCIENTE</a> | FunctionResult | Retorna a parte inteira de uma divisão |
| <a href="https://support.office.com/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">Função RADIANOS</a> | FunctionResult | Converte graus em radianos. |
| <a href="https://support.office.com/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">Função ALEATÓRIO</a> | FunctionResult | Retorna um número aleatório entre 0 e 1 |
| <a href="https://support.office.com/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">Função ALEATÓRIOENTRE</a> | FunctionResult | Retorna um número aleatório entre os números especificados |
| <a href="https://support.office.com/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">Função ORDEM.MÉD</a> | FunctionResult | Retorna a posição de um número em uma lista de números |
| <a href="https://support.office.com/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40" target="_blank">Função ORDEM.EQ</a> | FunctionResult | Retorna a posição de um número em uma lista de números |
| <a href="https://support.office.com/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">Função TAXA</a> | FunctionResult | Retorna a taxa de juros por período de uma anuidade |
| <a href="https://support.office.com/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">Função RECEBIDO</a> | FunctionResult | Retorna a quantia recebida no vencimento de um título totalmente investido |
| <a href="https://support.office.com/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a" target="_blank">Funções MUDAR, SUBSTITUIRB</a> | FunctionResult | Substitui caracteres em texto |
| <a href="https://support.office.com/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">Função REPT</a> | FunctionResult | Repete o texto um determinado número de vezes |
| <a href="https://support.office.com/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">Funções DIREITA, DIREITAB</a> | FunctionResult | Retorna os caracteres mais à direita de um valor de texto |
| <a href="https://support.office.com/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">Função ROMANO</a> | FunctionResult | Converte um algarismo arábico em romano, como texto |
| <a href="https://support.office.com/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">Função ARRED</a> | FunctionResult | Arredonda um número até uma quantidade especificada de dígitos |
| <a href="https://support.office.com/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">Função ARREDONDAR.PARA.BAIXO</a> | FunctionResult | Arredonda um número para baixo até zero |
| <a href="https://support.office.com/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">Função ARREDONDAR.PARA.CIMA</a> | FunctionResult | Arredonda um número para cima afastando-o de zero |
| <a href="https://support.office.com/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">Função LINS</a> | FunctionResult | Retorna o número de linhas em uma referência |
| <a href="https://support.office.com/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">Função TAXAJURO</a> | FunctionResult | Retorna uma taxa de juros equivalente para o crescimento de um investimento |
| <a href="https://support.office.com/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">Função SEC</a> | FunctionResult | Retorna a secante de um ângulo |
| <a href="https://support.office.com/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">Função SECH</a> | FunctionResult | Retorna a secante hiperbólica de um ângulo |
| <a href="https://support.office.com/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">Função SEGUNDO</a> | FunctionResult | Converte um número de série em um segundo |
| <a href="https://support.office.com/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">Função SOMASEQUÊNCIA</a> | FunctionResult | Retorna a soma de uma série polinomial baseada na fórmula |
| <a href="https://support.office.com/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">Função PLAN</a> | FunctionResult | Retorna o número da planilha referenciada |
| <a href="https://support.office.com/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">Função PLANS</a> | FunctionResult | Retorna o número de planilhas em uma referência |
| <a href="https://support.office.com/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">Função SINAL</a> | FunctionResult | Retorna o sinal de um número |
| <a href="https://support.office.com/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">Função SEN</a> | FunctionResult | Retorna o seno do ângulo fornecido |
| <a href="https://support.office.com/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">Função SENH</a> | FunctionResult | Retorna o seno hiperbólico de um número |
| <a href="https://support.office.com/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">Função DISTORÇÃO</a> | FunctionResult | Retorna a distorção de uma distribuição |
| <a href="https://support.office.com/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">Função DISTORÇÃO.P</a> | FunctionResult | Retorna a inclinação de uma distribuição com base em um preenchimento: uma caracterização do grau de assimetria de uma distribuição em torno de sua média |
| <a href="https://support.office.com/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">Função DPD</a> | FunctionResult | Retorna a depreciação em linha reta de um ativo durante um período |
| <a href="https://support.office.com/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07" target="_blank">Função MENOR</a> | FunctionResult | Retorna o menor valor k-ésimo em um conjunto de dados |
| <a href="https://support.office.com/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">Função RAIZ</a> | FunctionResult | Retorna uma raiz quadrada positiva |
| <a href="https://support.office.com/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">Função RAIZPI</a> | FunctionResult | Retorna a raiz quadrada de (número * pi) |
| <a href="https://support.office.com/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">Função PADRONIZAR</a> | FunctionResult | Retorna um valor normalizado |
| <a href="https://support.office.com/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">Função DESVPAD.P</a> | FunctionResult | Calcula o desvio padrão com base no preenchimento completo |
| <a href="https://support.office.com/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">Função DESVPAD.A</a> | FunctionResult | Estima o desvio padrão com base em uma amostra |
| <a href="https://support.office.com/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">Função DESVPADA</a> | FunctionResult | Estima o desvio padrão com base em uma amostra, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">Função DESVPADPA</a> | FunctionResult | Calcula o desvio padrão com base no preenchimento completo, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">Função SUBSTITUIR</a> | FunctionResult | Substitui um novo texto por um texto antigo em uma cadeia de caracteres de texto |
| <a href="https://support.office.com/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939" target="_blank">Função SUBTOTAL</a> | FunctionResult | Retorna um subtotal em uma lista ou banco de dados |
| <a href="https://support.office.com/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">Função SOMA</a> | FunctionResult | Adiciona os argumentos |
| <a href="https://support.office.com/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b" target="_blank">Função SOMASE</a> | FunctionResult | Adiciona as células especificadas por um determinado critério |
| <a href="https://support.office.com/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">Função SOMASES</a> | FunctionResult | Adiciona as células de um intervalo que atendam a vários critérios |
| <a href="https://support.office.com/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">Função SOMAQUAD</a> | FunctionResult | Retorna a soma dos quadrados dos argumentos. |
| <a href="https://support.office.com/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">Função SDA</a> | FunctionResult | Retorna a depreciação dos dígitos da soma dos anos de um ativo para um período especificado |
| <a href="https://support.office.com/article/T-function-fb83aeec-45e7-4924-af95-53e073541228" target="_blank">Função T</a> | FunctionResult | Converte os argumentos em texto |
| <a href="https://support.office.com/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">Função DIST.T</a> | FunctionResult | Retorna os Pontos Percentuais (probabilidade) para a distribuição t de Student |
| <a href="https://support.office.com/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">Função T.DIST.2T</a> | FunctionResult | Retorna os Pontos Percentuais (probabilidade) para a distribuição t de Student |
| <a href="https://support.office.com/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">Função DIST.T.CD</a> | FunctionResult | Retorna a distribuição t de Student |
| <a href="https://support.office.com/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">Função INV.T</a> | FunctionResult | Retorna o valor t da distribuição t de Student como uma função da probabilidade e dos graus de liberdade |
| <a href="https://support.office.com/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">Função T.INV.2T</a> | FunctionResult | Retorna o inverso da distribuição t de Student |
| <a href="https://support.office.com/article/TAN-function-08851a40-179f-4052-b789-d7f699447401" target="_blank">Função TAN</a> | FunctionResult | Retorna a tangente de um número |
| <a href="https://support.office.com/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">Função TANH</a> | FunctionResult | Retorna a tangente hiperbólica de um número |
| <a href="https://support.office.com/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">Função OTN</a> | FunctionResult | Retorna o rendimento de um título equivalente a uma obrigação do Tesouro |
| <a href="https://support.office.com/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">Função OTNVALOR</a> | FunctionResult | Retorna o preço por R$ 100,00 do valor nominal de uma obrigação do Tesouro |
| <a href="https://support.office.com/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">Função OTNLUCRO</a> | FunctionResult | Retorna o rendimento de uma obrigação do Tesouro |
| <a href="https://support.office.com/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">Função TEXTO</a> | FunctionResult | Formata um número e o converte em texto |
| <a href="https://support.office.com/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">Função TEMPO</a> | FunctionResult | Retorna o número de série de uma hora específica |
| <a href="https://support.office.com/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">Função VALOR.TEMPO</a> | FunctionResult | Converte um horário na forma de texto em um número de série |
| <a href="https://support.office.com/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">Função HOJE</a> | FunctionResult | Retorna o número de série da data de hoje |
| <a href="https://support.office.com/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">Função ARRUMAR</a> | FunctionResult | Remove espaços do texto |
| <a href="https://support.office.com/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">Função MÉDIA.INTERNA</a> | FunctionResult | Retorna a média do interior de um conjunto de dados |
| <a href="https://support.office.com/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">Função VERDADEIRO</a> | FunctionResult | Retorna o valor lógico `TRUE` |
| <a href="https://support.office.com/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">Função TRUNC</a> | FunctionResult | Trunca um número em um inteiro |
| <a href="https://support.office.com/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">Função TIPO</a> | FunctionResult | Retorna um número indicando o tipo de dados de um valor |
| <a href="https://support.office.com/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">Função CARACTUNICODE</a> | FunctionResult | Retorna o caractere Unicode referenciado por determinado valor numérico |
| <a href="https://support.office.com/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">Função UNICODE</a> | FunctionResult | Retorna o número (ponto de código) que corresponde ao primeiro caractere do texto |
| <a href="https://support.office.com/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">Função MAIÚSCULA</a> | FunctionResult | Converte texto em maiúsculas |
| <a href="https://support.office.com/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">Função VALOR</a> | FunctionResult | Converte um argumento de texto em um número |
| <a href="https://support.office.com/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">Função VAR.P</a> | FunctionResult | Calcula a variação com base no preenchimento inteiro |
| <a href="https://support.office.com/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b" target="_blank">Função VAR.A</a> | FunctionResult | Estima a variação com base em uma amostra |
| <a href="https://support.office.com/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">Função VARA</a> | FunctionResult | Estima a variação com base em uma amostra, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">Função VARPA</a> | FunctionResult | Calcula a variação com base no preenchimento total, incluindo números, texto e valores lógicos |
| <a href="https://support.office.com/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">Função BDV</a> | FunctionResult | Retorna a depreciação de um ativo para um período especificado ou parcial usando um método de balanço declinante |
| <a href="https://support.office.com/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">Função PROCV</a> | FunctionResult | Procura na primeira coluna de uma matriz e se move ao longo da linha para retornar o valor de uma célula |
| <a href="https://support.office.com/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">Função DIA.DA.SEMANA</a> | FunctionResult | Converte um número de série em um dia da semana |
| <a href="https://support.office.com/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">Função NÚMSEMANA</a> | FunctionResult | Converte um número de série em um número que representa onde a semana cai numericamente em um ano |
| <a href="https://support.office.com/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">Função DIST.WEIBULL</a> | FunctionResult | Retorna a distribuição de Weibull |
| <a href="https://support.office.com/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">Função DIATRABALHO</a> | FunctionResult | Retorna o número de série da data antes ou depois de um número específico de dias úteis |
| <a href="https://support.office.com/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">Função DIATRABALHO.INTL</a> | FunctionResult | Retorna o número de série da data antes ou depois de um número específico de dias úteis usando parâmetros para indicar quais e quantos dias são de fim de semana |
| <a href="https://support.office.com/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">Função XIRR</a> | FunctionResult | Fornece a taxa interna de retorno para um programa de fluxos de caixa que não é necessariamente periódico |
| <a href="https://support.office.com/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">Função XVPL</a> | FunctionResult | Retorna o valor presente líquido de um programa de fluxos de caixa que não é necessariamente periódico |
| <a href="https://support.office.com/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">Função XOR</a> | FunctionResult | Retorna um OU exclusivo lógico de todos os argumentos |
| <a href="https://support.office.com/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">Função ANO</a> | FunctionResult | Converte um número de série em um ano |
| <a href="https://support.office.com/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">Função FRAÇÃOANO</a> | FunctionResult | Retorna a fração do ano que representa o número de dias entre a data_inicial e a data_final |
| <a href="https://support.office.com/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">Função LUCRO</a> | FunctionResult | Retorna o lucro de um título que paga juros periódicos |
| <a href="https://support.office.com/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">Função LUCRODESC</a> | FunctionResult | Retorna o rendimento anual de um título descontado. Por exemplo, uma obrigação do Tesouro |
| <a href="https://support.office.com/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">Função LUCROVENC</a> | FunctionResult | Retorna o rendimento anual de um título que paga juros no vencimento |
| <a href="https://support.office.com/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Função TESTE.Z</a> | FunctionResult | Retorna o valor de probabilidade unicaudal do teste-z |

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Classe de funções (JavaScript API para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.functions)
- [Objeto Workbook de funções (JavaScript API para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.workbook#functions)
