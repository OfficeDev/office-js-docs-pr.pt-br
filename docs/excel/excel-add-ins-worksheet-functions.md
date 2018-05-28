---
title: Chamar fun??es internas de planilha do Excel usando as APIs JavaScript do Excel
description: ''
ms.date: 01/24/2017
ms.openlocfilehash: 5eb78484917cc3d4700c95d69fb1e83b6836d1cc
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="call-built-in-excel-worksheet-functions"></a>Chamar fun??es internas de planilha do Excel

Este artigo explica como chamar fun??es internas de planilha do Excel, como `VLOOKUP` e `SUM`, usando as API JavaScript do Excel. Tamb?m fornece a lista completa de fun??es internas de planilha Excel que podem ser chamadas usando as APIs JavaScript do Excel.

> [!NOTE]
> Para saber mais sobre como criar *fun??es personalizadas* no Excel usando as APIs JavaScript do Excel, confira [Criar fun??es personalizadas no Excel](custom-functions-overview.md).

## <a name="calling-a-worksheet-function"></a>Chamar uma fun??o de planilha

O trecho de c?digo a seguir mostra como chamar uma fun??o de planilha, onde `sampleFunction()` ? um espa?o reservado que deve ser substitu?do pelo nome da fun??o a chamar e os par?metros de entrada que a fun??o exige. A propriedade **valor** do objeto **FunctionResult** que uma fun??o de planilha retorna cont?m o resultado da fun??o especificada. Como mostra este exemplo, voc? deve carregar `load` a propriedade **valor** do objeto **FunctionResult** antes de l?-lo. Neste exemplo, o resultado da fun??o est? simplesmente sendo gravado no console. 

```js
var functionResult = context.workbook.functions.sampleFunction(); 
functionResult.load('value');
return context.sync()
    .then(function () {
        console.log('Result of the function: ' + functionResult.value);
    });
```

> [!TIP]
> Confira a se??o [Fun??es de planilha com suporte](#supported-worksheet-functions) deste artigo para obter uma lista das fun??es que podem ser chamadas usando as APIs JavaScript do Excel.

## <a name="sample-data"></a>Dados de exemplo

A imagem a seguir mostra uma tabela em uma planilha do Excel com dados de vendas para v?rios tipos de ferramentas durante um per?odo de tr?s meses. Cada n?mero da tabela representa o n?mero de unidades vendidas de uma ferramenta espec?fica em um m?s espec?fico. Os exemplos a seguir mostram como aplicar fun??es internas da planilha nesses dados.

![Captura de tela dos dados de vendas no Excel para martelo, chave inglesa e serra nos meses de novembro, dezembro e janeiro](../images/worksheet-functions-chaining-results.jpg)

## <a name="example-1-single-function"></a>Exemplo 1: fun??o individual

O exemplo a seguir se aplica ? fun??o `VLOOKUP` para os dados de exemplo descritos anteriormente a fim de identificar o n?mero de chaves inglesas vendidas em novembro.

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

## <a name="example-2-nested-functions"></a>Exemplo 2: fun??es aninhadas

O exemplo de c?digo a seguir aplica a fun??o `VLOOKUP` nos dados de amostras descritos anteriormente para identificar o n?mero de chaves inglesas vendidas em novembro e em dezembro e, em seguida, aplica a fun??o `SUM` para calcular o total de chaves inglesas vendido durante esses dois meses. 

Como mostra este exemplo, quando uma ou mais chamadas de fun??o s?o aninhadas dentro de outra chamada de fun??o, voc? s? precisa dar `load` no resultado final caso voc? queira ler (neste exemplo, `sumOfTwoLookups`). Os resultados intermedi?rios (neste exemplo, o resultado de cada fun??o `VLOOKUP`) ser?o calculados e usados para calcular o resultado final.

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

## <a name="supported-worksheet-functions"></a>Fun??es de planilha com suporte

As seguintes fun??es internas de planilhas do Excel podem ser chamadas usando as APIs JavaScript do Excel.

| Fun??o | Tipo de retorno | Descri??o |
|:---------------|:-------------|:-----------|
| <a href="https://support.office.com/en-us/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c" target="_blank">Fun??o ABS</a> | FunctionResult | Retorna o valor absoluto de um n?mero |
| <a href="https://support.office.com/en-us/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74" target="_blank">Fun??o JUROSACUM</a> | FunctionResult | Retorna juros acumulados de um t?tulo que paga juros peri?dicos |
| <a href="https://support.office.com/en-us/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7" target="_blank">Fun??o JUROSACUMV</a> | FunctionResult | Retorna juros acumulados de um t?tulo que paga juros no vencimento |
| <a href="https://support.office.com/en-us/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b" target="_blank">Fun??o ACOS</a> | FunctionResult | Retorna o arco cosseno de um n?mero |
| <a href="https://support.office.com/en-us/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe" target="_blank">Fun??o ACOSH</a> | FunctionResult | Retorna o cosseno hiperb?lico inverso de um n?mero |
| <a href="https://support.office.com/en-us/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905" target="_blank">Fun??o ACOT</a> | FunctionResult | Retorna o arco cotangente de um n?mero |
| <a href="https://support.office.com/en-us/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f" target="_blank">Fun??o ACOTH</a> | FunctionResult | Retorna o arco cotangente hiperb?lico de um n?mero |
| <a href="https://support.office.com/en-us/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51" target="_blank">Fun??o AMORDEGRC</a> | FunctionResult | Retorna a deprecia??o para cada per?odo cont?bil usando o coeficiente de deprecia??o |
| <a href="https://support.office.com/en-us/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8" target="_blank">Fun??o AMORLINC</a> | FunctionResult | Retorna a deprecia??o para cada per?odo cont?bil |
| <a href="https://support.office.com/en-us/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9" target="_blank">Fun??o E</a> | FunctionResult | Retorna `TRUE` se todos os seus argumentos forem verdadeiros |
| <a href="https://support.office.com/en-us/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f" target="_blank">Fun??o AR?BICO</a> | FunctionResult | Converte um n?mero romano em ar?bico, como um n?mero |
| <a href="https://support.office.com/en-us/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152" target="_blank">Fun??o ?REAS</a> | FunctionResult | Retorna o n?mero de ?reas em uma refer?ncia |
| <a href="https://support.office.com/en-us/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266" target="_blank">Fun??o ASC</a> | FunctionResult | Altera letras do ingl?s ou katakana de largura total (bytes duplos) dentro de uma cadeia de caracteres para caracteres de meia largura (byte ?nico) |
| <a href="https://support.office.com/en-us/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347" target="_blank">Fun??o ASEN</a> | FunctionResult | Retorna o arco seno de um n?mero |
| <a href="https://support.office.com/en-us/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c" target="_blank">Fun??o ASENH</a> | FunctionResult | Retorna o seno hiperb?lico inverso de um n?mero |
| <a href="https://support.office.com/en-us/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543" target="_blank">Fun??o ATAN</a> | FunctionResult | Retorna o arco tangente de um n?mero |
| <a href="https://support.office.com/en-us/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033" target="_blank">Fun??o ATAN2</a> | FunctionResult | Retorna o arco tangente das coordenadas x e y |
| <a href="https://support.office.com/en-us/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90" target="_blank">Fun??o ATANH</a> | FunctionResult | Retorna a tangente hiperb?lica inversa de um n?mero |
| <a href="https://support.office.com/en-us/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639" target="_blank">Fun??o DESV.M?DIO</a> | FunctionResult | Retorna a m?dia dos desvios absolutos dos pontos de dados a partir de sua m?dia |
| <a href="https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6" target="_blank">Fun??o M?DIA</a> | FunctionResult | Retorna a m?dia dos argumentos |
| <a href="https://support.office.com/en-us/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091" target="_blank">Fun??o M?DIAA</a> | FunctionResult | Retorna a m?dia dos argumentos, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642" target="_blank">Fun??o M?DIASE</a> | FunctionResult | Retorna a m?dia (m?dia aritm?tica) de todas as c?lulas em um intervalo que atendem a um determinado crit?rio |
| <a href="https://support.office.com/en-us/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690" target="_blank">Fun??o M?DIASES</a> | FunctionResult | Retorna a m?dia (m?dia aritm?tica) de todas as c?lulas que satisfazem v?rios crit?rios |
| <a href="https://support.office.com/en-us/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c" target="_blank">Fun??o BAHTTEXT</a> | FunctionResult | Converte um n?mero em texto, usando o formato de moeda ? (baht) |
| <a href="https://support.office.com/en-us/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342" target="_blank">Fun??o BASE</a> | FunctionResult | Converte um n?mero em uma representa??o de texto com a determinada base |
| <a href="https://support.office.com/en-us/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df" target="_blank">Fun??o BESSELI</a> | FunctionResult | Retorna a fun??o de Bessel In(x) modificada |
| <a href="https://support.office.com/en-us/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7" target="_blank">Fun??o BESSELJ</a> | FunctionResult | Retorna a fun??o de Bessel Jn(x) |
| <a href="https://support.office.com/en-us/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70" target="_blank">Fun??o BESSELK</a> | FunctionResult | Retorna a fun??o de Bessel Kn(x) modificada |
| <a href="https://support.office.com/en-us/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2" target="_blank">Fun??o BESSELY</a> | FunctionResult | Retorna a fun??o de Bessel Yn(x) |
| <a href="https://support.office.com/en-us/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31" target="_blank">Fun??o DIST.BETA</a> | FunctionResult | Retorna a fun??o de distribui??o cumulativa beta |
| <a href="https://support.office.com/en-us/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb" target="_blank">Fun??o INV.BETA</a> | FunctionResult | Retorna o inverso da fun??o de distribui??o cumulativa para uma distribui??o beta especificada |
| <a href="https://support.office.com/en-us/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c" target="_blank">Fun??o BIN2DEC</a> | FunctionResult | Converte um n?mero bin?rio em decimal |
| <a href="https://support.office.com/en-us/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc" target="_blank">Fun??o BIN2HEX</a> | FunctionResult | Converte um n?mero bin?rio em hexadecimal |
| <a href="https://support.office.com/en-us/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd" target="_blank">Fun??o BIN2OCT</a> | FunctionResult | Converte um n?mero bin?rio em octal |
| <a href="https://support.office.com/en-us/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c" target="_blank">Fun??o DISTR.BINOM</a> | FunctionResult | Retorna a probabilidade de distribui??o binomial do termo individual |
| <a href="https://support.office.com/en-us/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595" target="_blank">Fun??o INTERV.DISTR.BINOM</a> | FunctionResult | Retorna a probabilidade de um resultado de teste usando uma distribui??o binomial |
| <a href="https://support.office.com/en-us/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9" target="_blank">Fun??o INV.BINOM</a> | FunctionResult | Retorna o menor valor para o qual a distribui??o binomial cumulativa ? maior ou igual ao valor padr?o |
| <a href="https://support.office.com/en-us/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a" target="_blank">Fun??o BITAND</a> | FunctionResult | Retorna um bit "E" de dois n?meros |
| <a href="https://support.office.com/en-us/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c" target="_blank">Fun??o DESLOCESQBIT</a> | FunctionResult | Retorna um valor num?rico deslocado ? esquerda por quantidade_deslocamento bits |
| <a href="https://support.office.com/en-us/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2" target="_blank">Fun??o BITOR</a> | FunctionResult | Retorna um bit "OU" de dois n?meros |
| <a href="https://support.office.com/en-us/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c" target="_blank">Fun??o DESLOCDIRBIT</a> | FunctionResult | Retorna um valor num?rico deslocado ? direita por quantidade_deslocamento bits |
| <a href="https://support.office.com/en-us/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4" target="_blank">Fun??o BITXOR</a> | FunctionResult | Retorna um bit 'Exclusivo Ou' de dois n?meros |
| <a href="https://support.office.com/en-us/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8" target="_blank">Fun??o TETO.MAT</a> | FunctionResult | Arredonda um n?mero para cima, para o inteiro mais pr?ximo ou para o m?ltiplo mais pr?ximo significativo |
| <a href="https://support.office.com/en-us/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb" target="_blank">Fun??o TETO.PRECISO</a> | FunctionResult | Arredonda um n?mero para o inteiro mais pr?ximo ou para o m?ltiplo mais pr?ximo significativo. Independentemente do sinal do n?mero, ele ? arredondado para cima. |
| <a href="https://support.office.com/en-us/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a" target="_blank">Fun??o CARACT</a> | FunctionResult | Retorna o caractere especificado pelo n?mero de c?digo |
| <a href="https://support.office.com/en-us/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732" target="_blank">Fun??o DIST.QUIQUA</a> | FunctionResult | Retorna a fun??o de densidade da probabilidade beta cumulativa |
| <a href="https://support.office.com/en-us/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2" target="_blank">Fun??o DIST.QUIQUA.CD</a> | FunctionResult | Retorna a probabilidade unicaudal da distribui??o qui-quadrada |
| <a href="https://support.office.com/en-us/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f" target="_blank">Fun??o INV.QUIQUA</a> | FunctionResult | Retorna a fun??o de densidade da probabilidade beta cumulativa |
| <a href="https://support.office.com/en-us/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe" target="_blank">Fun??o INV.QUIQUA.CD</a> | FunctionResult | Retorna o inverso da probabilidade unicaudal da distribui??o qui-quadrada |
| <a href="https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc" target="_blank">Fun??o ESCOLHER</a> | FunctionResult | Escolhe um valor em uma lista de valores |
| <a href="https://support.office.com/en-us/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41" target="_blank">Fun??o TIRAR</a> | FunctionResult | Remove do texto todos os caracteres n?o imprim?veis |
| <a href="https://support.office.com/en-us/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928" target="_blank">Fun??o C?DIGO</a> | FunctionResult | Retorna um c?digo num?rico para o primeiro caractere de uma cadeia de texto |
| <a href="https://support.office.com/en-us/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca" target="_blank">Fun??o COLS</a> | FunctionResult | Retorna o n?mero de colunas em uma refer?ncia |
| <a href="https://support.office.com/en-us/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a" target="_blank">Fun??o COMBIN</a> | FunctionResult | Retorna o n?mero de combina??es de um determinado n?mero de objetos |
| <a href="https://support.office.com/en-us/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d" target="_blank">Fun??o COMBINA</a> | FunctionResult | Retorna o n?mero de combina??es com repeti??es de um determinado n?mero de itens |
| <a href="https://support.office.com/en-us/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128" target="_blank">Fun??o COMPLEXO</a> | FunctionResult | Converte coeficientes reais e imagin?rios em um n?mero complexo |
| <a href="https://support.office.com/en-us/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d" target="_blank">Fun??o CONCATENAR</a> | FunctionResult | Agrupa v?rios itens de texto em um item de texto |
| <a href="https://support.office.com/en-us/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4" target="_blank">Fun??o INT.CONFIAN?A.NORM</a> | FunctionResult | Retorna o intervalo de confian?a para um meio de preenchimento |
| <a href="https://support.office.com/en-us/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53" target="_blank">Fun??o INT.CONFIAN?A.T</a> | FunctionResult | Retorna o intervalo de confian?a para um meio de preenchimento, usando a distribui??o t de Student |
| <a href="https://support.office.com/en-us/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2" target="_blank">Fun??o CONVERTER</a> | FunctionResult | Converte um n?mero de um sistema de medidas para outro |
| <a href="https://support.office.com/en-us/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05" target="_blank">Fun??o COS</a> | FunctionResult | Retorna o cosseno de um n?mero |
| <a href="https://support.office.com/en-us/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555" target="_blank">Fun??o COSH</a> | FunctionResult | Retorna o cosseno hiperb?lico de um n?mero |
| <a href="https://support.office.com/en-us/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a" target="_blank">Fun??o COT</a> | FunctionResult | Retorna a cotangente de um ?ngulo |
| <a href="https://support.office.com/en-us/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5" target="_blank">Fun??o COTH</a> | FunctionResult | Retorna a cotangente hiperb?lica de um n?mero |
| <a href="https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c" target="_blank">Fun??o CONT.N?M</a> | FunctionResult | Calcula quantos n?meros h? na lista de argumentos |
| <a href="https://support.office.com/en-us/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509" target="_blank">Fun??o CONT.VALORES</a> | FunctionResult | Calcula quantos valores h? na lista de argumentos |
| <a href="https://support.office.com/en-us/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22" target="_blank">Fun??o CONTAR.VAZIO</a> | FunctionResult | Conta o n?mero de c?lulas vazias no intervalo especificado |
| <a href="https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34" target="_blank">Fun??o CONT.SE</a> | FunctionResult | Conta o n?mero de c?lulas em um intervalo que atendem aos crit?rios fornecidos |
| <a href="https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842" target="_blank">Fun??o CONT.SES</a> | FunctionResult | Conta o n?mero de c?lulas dentro de um intervalo que atende a m?ltiplos crit?rios |
| <a href="https://support.office.com/en-us/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872" target="_blank">Fun??o CUPDIASINLIQ</a> | FunctionResult | Retorna o n?mero de dias do in?cio do per?odo de cupom at? a data de liquida??o |
| <a href="https://support.office.com/en-us/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671" target="_blank">Fun??o CUPDIAS</a> | FunctionResult | Retorna o n?mero de dias no per?odo de cupom que cont?m a data de liquida??o |
| <a href="https://support.office.com/en-us/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547" target="_blank">Fun??o CUPDIASPR?X</a> | FunctionResult | Retorna o n?mero de dias da data de liquida??o at? a data do pr?ximo cupom |
| <a href="https://support.office.com/en-us/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f" target="_blank">Fun??o CUPDATAPR?X</a> | FunctionResult | Retorna a pr?xima data de cupom ap?s a data de quita??o |
| <a href="https://support.office.com/en-us/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522" target="_blank">Fun??o CUPN?M</a> | FunctionResult | Retorna o n?mero de cupons pag?veis entre as datas de quita??o e vencimento |
| <a href="https://support.office.com/en-us/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3" target="_blank">Fun??o CUPDATAANT</a> | FunctionResult | Retorna a data de cupom anterior ? data de quita??o |
| <a href="https://support.office.com/en-us/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1" target="_blank">Fun??o COSEC</a> | FunctionResult | Retorna a cossecante de um ?ngulo |
| <a href="https://support.office.com/en-us/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb" target="_blank">Fun??o COSECH</a> | FunctionResult | Retorna a cossecante hiperb?lica de um ?ngulo |
| <a href="https://support.office.com/en-us/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606" target="_blank">Fun??o PGTOJURACUM</a> | FunctionResult | Retorna os juros acumulados pagos entre dois per?odos |
| <a href="https://support.office.com/en-us/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d" target="_blank">Fun??o PGTOCAPACUM</a> | FunctionResult | Retorna o capital acumulado pago sobre um empr?stimo entre dois per?odos |
| <a href="https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349" target="_blank">Fun??o DATA</a> | FunctionResult | Retorna o n?mero de s?rie de uma data espec?fica |
| <a href="https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252" target="_blank">Fun??o DATA.VALOR</a> | FunctionResult | Converte uma data na forma de texto em um n?mero de s?rie |
| <a href="https://support.office.com/en-us/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee" target="_blank">Fun??o BDM?DIA</a> | FunctionResult | Retorna a m?dia das entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/en-us/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101" target="_blank">Fun??o DIA</a> | FunctionResult | Converte um n?mero de s?rie em um dia do m?s |
| <a href="https://support.office.com/en-us/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df" target="_blank">Fun??o DIAS</a> | FunctionResult | Retorna o n?mero de dias entre duas datas |
| <a href="https://support.office.com/en-us/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a" target="_blank">Fun??o DIAS360</a> | FunctionResult | Calcula o n?mero de dias entre duas datas com base em um ano de 360 dias |
| <a href="https://support.office.com/en-us/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7" target="_blank">Fun??o BD</a> | FunctionResult | Retorna a deprecia??o de um ativo para um per?odo especificado, usando o m?todo de balan?o de decl?nio fixo |
| <a href="https://support.office.com/en-us/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5" target="_blank">Fun??o DBCS</a> | FunctionResult | Altera letras do ingl?s ou katakana de meia largura (byte ?nico) dentro de uma cadeia de caracteres para caracteres de largura total (bytes duplos) |
| <a href="https://support.office.com/en-us/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1" target="_blank">Fun??o BDCONTAR</a> | FunctionResult | Conta as c?lulas que cont?m n?meros em um banco de dados |
| <a href="https://support.office.com/en-us/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244" target="_blank">Fun??o BDCONTARA</a> | FunctionResult | Conta c?lulas n?o vazias em um banco de dados |
| <a href="https://support.office.com/en-us/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5" target="_blank">Fun??o BDD</a> | FunctionResult | Retorna a deprecia??o de um ativo com rela??o a um per?odo especificado usando o m?todo de saldos decrescentes duplos ou qualquer outro m?todo especificado por voc? |
| <a href="https://support.office.com/en-us/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838" target="_blank">Fun??o DEC2BIN</a> | FunctionResult | Converte um n?mero decimal em bin?rio |
| <a href="https://support.office.com/en-us/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619" target="_blank">Fun??o DEC2HEX</a> | FunctionResult | Converte um n?mero decimal em hexadecimal |
| <a href="https://support.office.com/en-us/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f" target="_blank">Fun??o DEC2OCT</a> | FunctionResult | Converte um n?mero decimal em octal |
| <a href="https://support.office.com/en-us/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e" target="_blank">Fun??o DECIMAL</a> | FunctionResult | Converte em um n?mero decimal a representa??o de texto de um n?mero em determinada base |
| <a href="https://support.office.com/en-us/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1" target="_blank">Fun??o GRAUS</a> | FunctionResult | Converte radianos em graus |
| <a href="https://support.office.com/en-us/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432" target="_blank">Fun??o DELTA</a> | FunctionResult | Testa se dois valores s?o iguais |
| <a href="https://support.office.com/en-us/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444" target="_blank">Fun??o DESVQ</a> | FunctionResult | Retorna a soma dos quadrados dos desvios |
| <a href="https://support.office.com/en-us/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e" target="_blank">Fun??o BDEXTRAIR</a> | FunctionResult | Extrai de um banco de dados um ?nico registro que corresponde aos crit?rios especificados |
| <a href="https://support.office.com/en-us/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53" target="_blank">Fun??o DESC</a> | FunctionResult | Retorna a taxa de desconto de um t?tulo |
| <a href="https://support.office.com/en-us/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2" target="_blank">Fun??o BDM?X</a> | FunctionResult | Retorna o valor m?ximo de entradas selecionadas de banco de dados |
| <a href="https://support.office.com/en-us/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3" target="_blank">Fun??o BDM?N</a> | FunctionResult | Retorna o valor m?nimo de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/en-us/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611" target="_blank">Fun??o MOEDA</a> | FunctionResult | Converte um n?mero em texto, usando o formato de moeda $ (cifr?o) |
| <a href="https://support.office.com/en-us/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427" target="_blank">Fun??o MOEDADEC</a> | FunctionResult | Converte um pre?o em moeda expresso como uma fra??o em um pre?o em moeda expresso como um n?mero decimal |
| <a href="https://support.office.com/en-us/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495" target="_blank">Fun??o MOEDAFRA</a> | FunctionResult | Converte um pre?o em moeda expresso como um n?mero decimal em um pre?o em moeda expresso como uma fra??o |
| <a href="https://support.office.com/en-us/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31" target="_blank">Fun??o BDMULTIPL</a> | FunctionResult | Multiplica os valores em um campo espec?fico de registros que correspondem ao crit?rio em um banco de dados |
| <a href="https://support.office.com/en-us/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96" target="_blank">Fun??o BDEST</a> | FunctionResult | Estima o desvio padr?o com base em uma amostra de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/en-us/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b" target="_blank">Fun??o BDDESVPA</a> | FunctionResult | Calcula o desvio padr?o com base no preenchimento completo de entradas selecionadas de banco de dados |
| <a href="https://support.office.com/en-us/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b" target="_blank">Fun??o BDSOMA</a> | FunctionResult | Soma os n?meros na coluna de campos de registros do banco de dados que correspondem ao crit?rio |
| <a href="https://support.office.com/en-us/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038" target="_blank">Fun??o DURA??O</a> | FunctionResult | Retorna a dura??o anual de um t?tulo com pagamentos de juros peri?dicos |
| <a href="https://support.office.com/en-us/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1" target="_blank">Fun??o BDVAREST</a> | FunctionResult | Estima a varia??o com base em uma amostra de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/en-us/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc" target="_blank">Fun??o BDVARP</a> | FunctionResult | Calcula a varia??o com base no preenchimento completo de entradas selecionadas de um banco de dados |
| <a href="https://support.office.com/en-us/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5" target="_blank">Fun??o DATAM</a> | FunctionResult | Retorna o n?mero de s?rie da data que ? o n?mero indicado de meses antes ou depois da data inicial |
| <a href="https://support.office.com/en-us/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4" target="_blank">Fun??o EFETIVA</a> | FunctionResult | Retorna a taxa de juros anual efetiva |
| <a href="https://support.office.com/en-us/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628" target="_blank">Fun??o FIMM?S</a> | FunctionResult | Retorna o n?mero de s?rie do ?ltimo dia do m?s antes ou depois de um n?mero especificado de meses |
| <a href="https://support.office.com/en-us/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349" target="_blank">Fun??o FUNERRO</a> | FunctionResult | Retorna a fun??o de erro |
| <a href="https://support.office.com/en-us/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0" target="_blank">Fun??o FUNERRO.PRECISO</a> | FunctionResult | Retorna a fun??o de erro |
| <a href="https://support.office.com/en-us/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3" target="_blank">Fun??o FUNERROCOMPL</a> | FunctionResult | Retorna a fun??o de erro complementar |
| <a href="https://support.office.com/en-us/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273" target="_blank">Fun??o FUNERROCOMPL.PRECISO</a> | FunctionResult | Retorna a fun??o FUNERRO complementar integrada entre x e infinito |
| <a href="https://support.office.com/en-us/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa" target="_blank">Fun??o TIPO.ERRO</a> | FunctionResult | Retorna um n?mero correspondente a um tipo de erro |
| <a href="https://support.office.com/en-us/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9" target="_blank">Fun??o PAR</a> | FunctionResult | Arredonda um n?mero para cima at? o inteiro par mais pr?ximo |
| <a href="https://support.office.com/en-us/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926" target="_blank">Fun??o EXATO</a> | FunctionResult | Verifica se dois valores de texto s?o id?nticos |
| <a href="https://support.office.com/en-us/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe" target="_blank">Fun??o EXP</a> | FunctionResult | Retorna e elevado ? pot?ncia de um n?mero especificado |
| <a href="https://support.office.com/en-us/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e" target="_blank">Fun??o DISTR.EXPON</a> | FunctionResult | Retorna a distribui??o exponencial |
| <a href="https://support.office.com/en-us/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d" target="_blank">Fun??o DIST.F</a> | FunctionResult | Retorna a distribui??o de probabilidade F |
| <a href="https://support.office.com/en-us/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520" target="_blank">Fun??o DIST.F.CD</a> | FunctionResult | Retorna a distribui??o de probabilidade F |
| <a href="https://support.office.com/en-us/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe" target="_blank">Fun??o INV.F</a> | FunctionResult | Retorna o inverso da distribui??o de probabilidade F |
| <a href="https://support.office.com/en-us/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00" target="_blank">Fun??o INV.F.CD</a> | FunctionResult | Retorna o inverso da distribui??o de probabilidade F |
| <a href="https://support.office.com/en-us/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3" target="_blank">Fun??o FATORIAL</a> | FunctionResult | Retorna o fatorial de um n?mero |
| <a href="https://support.office.com/en-us/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8" target="_blank">Fun??o FATDUPLO</a> | FunctionResult | Retorna o fatorial duplo de um n?mero |
| <a href="https://support.office.com/en-us/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904" target="_blank">Fun??o FALSO</a> | FunctionResult | Retorna o valor l?gico `FALSE` |
| <a href="https://support.office.com/en-us/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628" target="_blank">Fun??es PROCURAR, PROCURARB</a> | FunctionResult | Procura um valor de texto dentro de outro (diferencia mai?sculas de min?sculas) |
| <a href="https://support.office.com/en-us/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69" target="_blank">Fun??o FISHER</a> | FunctionResult | Retorna a transforma??o Fisher |
| <a href="https://support.office.com/en-us/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb" target="_blank">Fun??o FISHERINV</a> | FunctionResult | Retorna o inverso da transforma??o Fisher |
| <a href="https://support.office.com/en-us/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a" target="_blank">Fun??o FIXO</a> | FunctionResult | Formata um n?mero como texto com um n?mero fixo de decimais |
| <a href="https://support.office.com/en-us/article/FLOOR-function-14bb497c-24f2-4e04-b327-b0b4de5a8886" target="_blank">Fun??o ARREDMULTB</a> | FunctionResult | Arredonda um n?mero para baixo at? zero |
| <a href="https://support.office.com/en-us/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5" target="_blank">Fun??o de ARREDMULTB.MAT</a> | FunctionResult | Arredonda um n?mero para baixo para o inteiro mais pr?ximo ou para o m?ltiplo mais pr?ximo de signific?ncia |
| <a href="https://support.office.com/en-us/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e" target="_blank">Fun??o ARREDMULTB.PRECISO</a> | FunctionResult | Arredonda um n?mero para baixo para o inteiro mais pr?ximo ou para o m?ltiplo mais pr?ximo de signific?ncia. Independentemente do sinal do n?mero, ele ? arredondado para baixo. |
| <a href="https://support.office.com/en-us/article/FORECAST-function-50ca49c9-7b40-4892-94e4-7ad38bbeda99" target="_blank">Fun??o PREVIS?O</a> | FunctionResult | Retorna um valor ao longo de uma tend?ncia linear |
| <a href="https://support.office.com/en-us/article/FORECASTETS-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">Fun??o PREVIS?O.ETS</a> | FunctionResult | Retorna um valor futuro com base em valores existentes (hist?ricos), usando a vers?o AAA do algoritmo de Suaviza??o Exponencial (ETS) |
| <a href="https://support.office.com/en-us/article/FORECASTETSCONFINT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">Fun??o PREVIS?O.ETS.CONFINT</a> | FunctionResult | Retorna um intervalo de confian?a para o valor de previs?o na data de destino especificada |
| <a href="https://support.office.com/en-us/article/FORECASTETSSEASONALITY-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">Fun??o PREVIS?O.ETS.SAZONALIDADE</a> | FunctionResult | Retorna o comprimento do padr?o repetitivo que Excel detecta para a s?rie temporal especificada |
| <a href="https://support.office.com/en-us/article/FORECASTETSSTAT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">Fun??o PREVIS?O.ETS.STAT</a> | FunctionResult | Retorna um valor estat?stico como resultado de previs?o de s?rie temporal |
| <a href="https://support.office.com/en-us/article/FORECASTLINEAR-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e" target="_blank">Fun??o PREVIS?O.LINEAR</a> | FunctionResult | Retorna um valor futuro com base em valores existentes |
| <a href="https://support.office.com/en-us/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3" target="_blank">Fun??o VF</a> | FunctionResult | Retorna o valor futuro de um investimento |
| <a href="https://support.office.com/en-us/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d" target="_blank">Fun??o VFPLANO</a> | FunctionResult | Retorna o valor futuro de um capital inicial ap?s a aplica??o de uma s?rie de taxas de juros compostas |
| <a href="https://support.office.com/en-us/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297" target="_blank">Fun??o GAMA</a> | FunctionResult | Retorna o valor da fun??o GAMA |
| <a href="https://support.office.com/en-us/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def" target="_blank">Fun??o DIST.GAMA</a> | FunctionResult | Retorna a distribui??o gama |
| <a href="https://support.office.com/en-us/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18" target="_blank">Fun??o INV.GAMA</a> | FunctionResult | Retorna o inverso da distribui??o cumulativa gama |
| <a href="https://support.office.com/en-us/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9" target="_blank">Fun??o LNGAMA</a> | FunctionResult | Retorna o logaritmo natural da fun??o gama, G(x) |
| <a href="https://support.office.com/en-us/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599" target="_blank">Fun??o LNGAMA.PRECISO</a> | FunctionResult | Retorna o logaritmo natural da fun??o gama, G(x) |
| <a href="https://support.office.com/en-us/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33" target="_blank">Fun??o GAUSS</a> | FunctionResult | Retorna menos 0,5 que a distribui??o cumulativa normal padr?o |
| <a href="https://support.office.com/en-us/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a" target="_blank">Fun??o MDC</a> | FunctionResult | Retorna o m?ximo divisor comum |
| <a href="https://support.office.com/en-us/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5" target="_blank">Fun??o M?DIA.GEOM?TRICA</a> | FunctionResult | Retorna a m?dia geom?trica |
| <a href="https://support.office.com/en-us/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df" target="_blank">Fun??o DEGRAU</a> | FunctionResult | Testa se um n?mero ? maior do que um valor limite |
| <a href="https://support.office.com/en-us/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6" target="_blank">Fun??o M?DIA.HARM?NICA</a> | FunctionResult | Retorna a m?dia harm?nica |
| <a href="https://support.office.com/en-us/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1" target="_blank">Fun??o HEX2BIN</a> | FunctionResult | Converte um n?mero hexadecimal em bin?rio |
| <a href="https://support.office.com/en-us/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e" target="_blank">Fun??o HEX2DEC</a> | FunctionResult | Converte um n?mero hexadecimal em decimal |
| <a href="https://support.office.com/en-us/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912" target="_blank">Fun??o HEX2OCT</a> | FunctionResult | Converte um n?mero hexadecimal em octal |
| <a href="https://support.office.com/en-us/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f" target="_blank">Fun??o PROCH</a> | FunctionResult | Procura na linha superior de uma matriz e retorna o valor da c?lula especificada |
| <a href="https://support.office.com/en-us/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7" target="_blank">Fun??o HORA</a> | FunctionResult | Converte um n?mero de s?rie em um hora |
| <a href="https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f" target="_blank">Fun??o HIPERLINK</a> | FunctionResult | Cria um atalho ou salto que abre um documento armazenado em um servidor de rede, uma intranet ou na Internet |
| <a href="https://support.office.com/en-us/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf" target="_blank">Fun??o DIST.HIPERGEOM.N</a> | FunctionResult | Retorna a distribui??o hipergeom?trica |
| <a href="https://support.office.com/en-us/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2" target="_blank">Fun??o SE</a> | FunctionResult | Especifica um teste l?gico a ser executado |
| <a href="https://support.office.com/en-us/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1" target="_blank">Fun??o IMABS</a> | FunctionResult | Retorna o valor absoluto (m?dulo) de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a" target="_blank">Fun??o IMAGIN?RIO</a> | FunctionResult | Retorna o coeficiente imagin?rio de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a" target="_blank">Fun??o IMARG</a> | FunctionResult | Retorna o argumento teta, um ?ngulo expresso em radianos |
| <a href="https://support.office.com/en-us/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42" target="_blank">Fun??o IMCONJ</a> | FunctionResult | Retorna o conjugado complexo de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c" target="_blank">Fun??o IMCOS</a> | FunctionResult | Retorna o cosseno de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff" target="_blank">Fun??o IMCOSH</a> | FunctionResult | Retorna o cosseno hiperb?lico de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c" target="_blank">Fun??o IMCOT</a> | FunctionResult | Retorna a cotangente de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323" target="_blank">Fun??o IMCOSEC</a> | FunctionResult | Retorna a cossecante de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9" target="_blank">Fun??o IMCOSECH</a> | FunctionResult | Retorna a cossecante hiperb?lica de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f" target="_blank">Fun??o IMDIV</a> | FunctionResult | Retorna o quociente de dois n?meros complexos |
| <a href="https://support.office.com/en-us/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f" target="_blank">Fun??o IMEXP</a> | FunctionResult | Retorna o exponencial de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8" target="_blank">Fun??o IMLN</a> | FunctionResult | Retorna o logaritmo natural de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5" target="_blank">Fun??o IMLOG10</a> | FunctionResult | Retorna o logaritmo de base 10 de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51" target="_blank">Fun??o IMLOG2</a> | FunctionResult | Retorna o logaritmo de base 2 de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39" target="_blank">Fun??o IMPOT</a> | FunctionResult | Retorna um n?mero complexo elevado a uma pot?ncia inteira |
| <a href="https://support.office.com/en-us/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba" target="_blank">Fun??o IMPROD</a> | FunctionResult | Retorna o produto de 2 a 255 n?meros complexos |
| <a href="https://support.office.com/en-us/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366" target="_blank">Fun??o IMREAL</a> | FunctionResult | Retorna o coeficiente real de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0" target="_blank">Fun??o IMSEC</a> | FunctionResult | Retorna a secante de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b" target="_blank">Fun??o IMSECH</a> | FunctionResult | Retorna a secante hiperb?lica de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6" target="_blank">Fun??o IMSENO</a> | FunctionResult | Retorna o seno de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d" target="_blank">Fun??o IMSENH</a> | FunctionResult | Retorna o seno hiperb?lico de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70" target="_blank">Fun??o IMSQRT</a> | FunctionResult | Retorna a raiz quadrada de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054" target="_blank">Fun??o IMSUBTR</a> | FunctionResult | Retorna a diferen?a entre dois n?meros complexos |
| <a href="https://support.office.com/en-us/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f" target="_blank">Fun??o IMSOMA</a> | FunctionResult | Retorna a soma de n?meros complexos |
| <a href="https://support.office.com/en-us/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132" target="_blank">Fun??o IMTAN</a> | FunctionResult | Retorna a tangente de um n?mero complexo |
| <a href="https://support.office.com/en-us/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef" target="_blank">Fun??o INT</a> | FunctionResult | Arredonda um n?mero para baixo at? o n?mero inteiro mais pr?ximo |
| <a href="https://support.office.com/en-us/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f" target="_blank">Fun??o TAXAJUROS</a> | FunctionResult | Retorna a taxa de juros de um t?tulo totalmente investido |
| <a href="https://support.office.com/en-us/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f" target="_blank">Fun??o IPGTO</a> | FunctionResult | Retorna o pagamento de juros para um investimento em um determinado per?odo |
| <a href="https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc" target="_blank">Fun??o TIR</a> | FunctionResult | Retorna a taxa interna de retorno de uma s?rie de fluxos de caixa |
| <a href="https://support.office.com/en-us/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?ERRO</a> | FunctionResult | Retorna `TRUE` se o valor for um valor de erro diferente de #N/D |
| <a href="https://support.office.com/en-us/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?ERROS</a> | FunctionResult | Retorna `TRUE` se o valor for um valor de erro |
| <a href="https://support.office.com/en-us/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356" target="_blank">Fun??o ?PAR</a> | FunctionResult | Retorna `TRUE` se o n?mero for par |
| <a href="https://support.office.com/en-us/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5" target="_blank">Fun??o ?F?RMULA</a> | FunctionResult | Retorna `TRUE` quando h? uma refer?ncia a uma c?lula que cont?m uma f?rmula |
| <a href="https://support.office.com/en-us/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?L?GICO</a> | FunctionResult | Retorna `TRUE` se o valor for um valor l?gico |
| <a href="https://support.office.com/en-us/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?.N?O.DISP</a> | FunctionResult | Retorna `TRUE` se o valor for o valor de erro #N/D |
| <a href="https://support.office.com/en-us/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?.N?O.TEXTO</a> | FunctionResult | Retorna `TRUE` se o valor for diferente de texto |
| <a href="https://support.office.com/en-us/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?N?M</a> | FunctionResult | Retorna `TRUE` se o valor for um n?mero |
| <a href="https://support.office.com/en-us/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83" target="_blank">Fun??o ISO.TETO</a> | FunctionResult | Retorna um n?mero para o inteiro mais pr?ximo ou para o m?ltiplo mais pr?ximo de signific?ncia |
| <a href="https://support.office.com/en-us/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?IMPAR</a> | FunctionResult | Retorna `TRUE` se o n?mero for ?mpar |
| <a href="https://support.office.com/en-us/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e" target="_blank">Fun??o N?MSEMANAISO</a> | FunctionResult | Retorna o n?mero do n?mero da semana ISO do ano referente a determinada data |
| <a href="https://support.office.com/en-us/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc" target="_blank">Fun??o ?PGTO</a> | FunctionResult | Calcula os juros pagos durante um per?odo espec?fico de um investimento |
| <a href="https://support.office.com/en-us/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?REF</a> | FunctionResult | Retorna `TRUE` se o valor for uma refer?ncia |
| <a href="https://support.office.com/en-us/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665" target="_blank">Fun??o ?TEXTO</a> | FunctionResult | Retorna `TRUE` se o valor for texto |
| <a href="https://support.office.com/en-us/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab" target="_blank">Fun??o CURT</a> | FunctionResult | Retorna a curtose de um conjunto de dados |
| <a href="https://support.office.com/en-us/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64" target="_blank">Fun??o MAIOR</a> | FunctionResult | Retorna o maior valor k-?simo em um conjunto de dados |
| <a href="https://support.office.com/en-us/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94" target="_blank">Fun??o MMC</a> | FunctionResult | Retorna o m?nimo m?ltiplo comum |
| <a href="https://support.office.com/en-us/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c" target="_blank">Fun??es ESQUERDA, ESQUERDAB</a> | FunctionResult | Retorna os caracteres mais ? esquerda de um valor de texto |
| <a href="https://support.office.com/en-us/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb" target="_blank">Fun??es N?M.CARACT, N?M.CARACTB</a> | FunctionResult | Retorna o n?mero de caracteres em uma cadeia de texto |
| <a href="https://support.office.com/en-us/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f" target="_blank">Fun??o LN</a> | FunctionResult | Retorna o logaritmo natural de um n?mero |
| <a href="https://support.office.com/en-us/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280" target="_blank">Fun??o LOG</a> | FunctionResult | Retorna o logaritmo de um n?mero de uma base especificada |
| <a href="https://support.office.com/en-us/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211" target="_blank">Fun??o LOG10</a> | FunctionResult | Retorna o logaritmo de base 10 de um n?mero |
| <a href="https://support.office.com/en-us/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070" target="_blank">Fun??o DIST.LOGNORMAL.N</a> | FunctionResult | Retorna a distribui??o lognormal cumulativa |
| <a href="https://support.office.com/en-us/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600" target="_blank">Fun??o INV.LOGNORMAL</a> | FunctionResult | Retorna o inverso da distribui??o cumulativa lognormal |
| <a href="https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb" target="_blank">Fun??o PROC</a> | FunctionResult | Procura valores em um vetor ou uma matriz |
| <a href="https://support.office.com/en-us/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4" target="_blank">Fun??o MIN?SCULA</a> | FunctionResult | Converte texto em min?sculas |
| <a href="https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a" target="_blank">Fun??o CORRESP</a> | FunctionResult | Procura valores em uma refer?ncia ou matriz |
| <a href="https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098" target="_blank">Fun??o M?XIMO</a> | FunctionResult | Retorna o valor m?ximo em uma lista de argumentos |
| <a href="https://support.office.com/en-us/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d" target="_blank">Fun??o M?XIMOA</a> | FunctionResult | Retorna o maior valor em uma lista de argumentos, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c" target="_blank">Fun??o MDURA??O</a> | FunctionResult | Retorna a dura??o modificada Macauley de um t?tulo com um valor de paridade equivalente a R$ 100 |
| <a href="https://support.office.com/en-us/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2" target="_blank">Fun??o MED</a> | FunctionResult | Retorna a mediana dos n?meros indicados |
| <a href="https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028" target="_blank">Fun??es EXT.TEXTO, EXT.TEXTOB</a> | FunctionResult | Retorna um n?mero espec?fico de caracteres de uma cadeia de texto come?ando na posi??o especificada |
| <a href="https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152" target="_blank">Fun??o M?NIMO</a> | FunctionResult | Retorna o valor m?nimo em uma lista de argumentos |
| <a href="https://support.office.com/en-us/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3" target="_blank">Fun??o M?NIMOA</a> | FunctionResult | Retorna o menor valor em uma lista de argumentos, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589" target="_blank">Fun??o MINUTO</a> | FunctionResult | Converte um n?mero de s?rie em um minuto |
| <a href="https://support.office.com/en-us/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524" target="_blank">Fun??o MTIR</a> | FunctionResult | Calcula a taxa interna de retorno em que fluxos de caixa positivos e negativos s?o financiados com diferentes taxas |
| <a href="https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3" target="_blank">Fun??o MOD</a> | FunctionResult | Retorna o resto da divis?o |
| <a href="https://support.office.com/en-us/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8" target="_blank">Fun??o M?S</a> | FunctionResult | Converte um n?mero de s?rie em um m?s |
| <a href="https://support.office.com/en-us/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427" target="_blank">Fun??o MARRED</a> | FunctionResult | Retorna um n?mero arredondado ao m?ltiplo desejado |
| <a href="https://support.office.com/en-us/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6" target="_blank">Fun??o MULTINOMIAL</a> | FunctionResult | Retorna o multin?mio de um conjunto de n?meros |
| <a href="https://support.office.com/en-us/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9" target="_blank">Fun??o N</a> | FunctionResult | Retorna um valor convertido em um n?mero |
| <a href="https://support.office.com/en-us/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c" target="_blank">Fun??o N?O.DISP</a> | FunctionResult | Retorna o valor de erro #N/D |
| <a href="https://support.office.com/en-us/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599" target="_blank">Fun??o DIST.BIN.NEG.N</a> | FunctionResult | Retorna a distribui??o binomial negativa |
| <a href="https://support.office.com/en-us/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7" target="_blank">Fun??o DIATRABALHOTOTAL</a> | FunctionResult | Retorna o n?mero de dias ?teis inteiros entre duas datas |
| <a href="https://support.office.com/en-us/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28" target="_blank">Fun??o DIATRABALHOTOTAL.INTL</a> | FunctionResult | Retorna o n?mero de dias de trabalho totais entre duas datas usando par?metros para indicar quais e quantos dias caem em finais de semana |
| <a href="https://support.office.com/en-us/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b" target="_blank">Fun??o NOMINAL</a> | FunctionResult | Retorna a taxa de juros nominal anual |
| <a href="https://support.office.com/en-us/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d" target="_blank">Fun??o DIST.NORM.N</a> | FunctionResult | Retorna a distribui??o cumulativa normal |
| <a href="https://support.office.com/en-us/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13" target="_blank">Fun??o INV.NORM.N</a> | FunctionResult | Retorna o inverso da distribui??o cumulativa normal |
| <a href="https://support.office.com/en-us/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88" target="_blank">Fun??o DIST.NORMP.N</a> | FunctionResult | Retorna a distribui??o cumulativa normal padr?o |
| <a href="https://support.office.com/en-us/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1" target="_blank">Fun??o INV.NORMP.N</a> | FunctionResult | Retorna o inverso da distribui??o cumulativa normal padr?o |
| <a href="https://support.office.com/en-us/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77" target="_blank">Fun??o N?O</a> | FunctionResult | Inverte o valor l?gico do argumento |
| <a href="https://support.office.com/en-us/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46" target="_blank">Fun??o AGORA</a> | FunctionResult | Retorna o n?mero de s?rie sequencial da data e hora atuais |
| <a href="https://support.office.com/en-us/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815" target="_blank">Fun??o NPER</a> | FunctionResult | Retorna o n?mero de per?odos de um investimento |
| <a href="https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568" target="_blank">Fun??o VPL</a> | FunctionResult | Retorna o valor l?quido atual de um investimento com base em uma s?rie de fluxos de caixa peri?dicos e em uma taxa de desconto |
| <a href="https://support.office.com/en-us/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879" target="_blank">Fun??o VALORNUM?RICO</a> | FunctionResult | Converte texto em n?mero de maneira independente de localidade |
| <a href="https://support.office.com/en-us/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589" target="_blank">Fun??o OCT2BIN</a> | FunctionResult | Converte um n?mero octal em bin?rio |
| <a href="https://support.office.com/en-us/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554" target="_blank">Fun??o OCT2DEC</a> | FunctionResult | Converte um n?mero octal em decimal |
| <a href="https://support.office.com/en-us/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f" target="_blank">Fun??o OCT2HEX</a> | FunctionResult | Converte um n?mero octal em hexadecimal |
| <a href="https://support.office.com/en-us/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98" target="_blank">Fun??o ?MPAR</a> | FunctionResult | Arredonda um n?mero para cima at? o inteiro ?mpar mais pr?ximo |
| <a href="https://support.office.com/en-us/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1" target="_blank">Fun??o PRE?OPRIMINC</a> | FunctionResult | Retorna o pre?o por R$ 100 do valor nominal de um t?tulo com um per?odo inicial incompleto |
| <a href="https://support.office.com/en-us/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37" target="_blank">Fun??o LUCROPRIMINC</a> | FunctionResult | Retorna o rendimento de um t?tulo com um per?odo inicial incompleto |
| <a href="https://support.office.com/en-us/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4" target="_blank">Fun??o PRE?O?LTINC</a> | FunctionResult | Retorna o pre?o por R$ 100 do valor nominal de um t?tulo com um per?odo final incompleto |
| <a href="https://support.office.com/en-us/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238" target="_blank">Fun??o LUCRO?LTINC</a> | FunctionResult | Retorna o rendimento de um t?tulo com um per?odo final incompleto |
| <a href="https://support.office.com/en-us/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0" target="_blank">Fun??o OU</a> | FunctionResult | Retorna `TRUE` se um dos argumentos for verdadeiro |
| <a href="https://support.office.com/en-us/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf" target="_blank">Fun??o DURA??OP</a> | FunctionResult | Retorna o n?mero de per?odos necess?rios para que um investimento atinja um valor espec?fico |
| <a href="https://support.office.com/en-us/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba" target="_blank">Fun??o PERCENTIL.EXC</a> | FunctionResult | Retorna o k-?simo percentil de valores em um intervalo, onde k est? no intervalo 0..1, exclusive |
| <a href="https://support.office.com/en-us/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed" target="_blank">Fun??o PERCENTIL.INC</a> | FunctionResult | Retorna o k-?simo percentil de valores em um intervalo |
| <a href="https://support.office.com/en-us/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314" target="_blank">Fun??o ORDEM.PORCENTUAL.EXC</a> | FunctionResult | Retorna a posi??o de um valor em um conjunto de dados como uma porcentagem (0..1, exclusivo) do conjunto de dados |
| <a href="https://support.office.com/en-us/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a" target="_blank">Fun??o ORDEM.PORCENTUAL.INC</a> | FunctionResult | Retorna a ordem percentual de um valor em um conjunto de dados |
| <a href="https://support.office.com/en-us/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3" target="_blank">Fun??o PERMUT</a> | FunctionResult | Retorna o n?mero de permuta??es de um determinado n?mero de objetos |
| <a href="https://support.office.com/en-us/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e" target="_blank">Fun??o PERMUTAS</a> | FunctionResult | Retorna o n?mero de permuta??es referentes a determinado n?mero de objetos (com repeti??es) que podem ser selecionadas do total de objetos |
| <a href="https://support.office.com/en-us/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c" target="_blank">Fun??o PHI</a> | FunctionResult | Retorna o valor da fun??o de densidade referente a uma distribui??o normal padr?o |
| <a href="https://support.office.com/en-us/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b" target="_blank">Fun??o PI</a> | FunctionResult | Retorna o valor de pi |
| <a href="https://support.office.com/en-us/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441" target="_blank">Fun??o PGTO</a> | FunctionResult | Retorna o pagamento peri?dico de uma anuidade |
| <a href="https://support.office.com/en-us/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636" target="_blank">Fun??o DIST.POISSON</a> | FunctionResult | Retorna a distribui??o de Poisson |
| <a href="https://support.office.com/en-us/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a" target="_blank">Fun??o POT?NCIA</a> | FunctionResult | Retorna o resultado de um n?mero elevado a uma pot?ncia |
| <a href="https://support.office.com/en-us/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b" target="_blank">Fun??o PPGTO</a> | FunctionResult | Retorna o pagamento de capital para determinado per?odo de investimento |
| <a href="https://support.office.com/en-us/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a" target="_blank">Fun??o PRE?O</a> | FunctionResult | Retorna o pre?o pelo valor nominal R$100 de um t?tulo que paga juros peri?dicos |
| <a href="https://support.office.com/en-us/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3" target="_blank">Fun??o PRE?ODESC</a> | FunctionResult | Retorna o pre?o por valor nominal de R$ 100,00 de um t?tulo descontado |
| <a href="https://support.office.com/en-us/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77" target="_blank">Fun??o PRE?OVENC</a> | FunctionResult | Retorna o pre?o pelo valor nominal R$100 de um t?tulo que paga juros no vencimento |
| <a href="https://support.office.com/en-us/article/PROB-function-9ac30561-c81c-4259-8253-34f0a238fc49" target="_blank">Fun??o PROB</a> | FunctionResult | Retorna a probabilidade de valores em um intervalo estarem entre dois limites |
| <a href="https://support.office.com/en-us/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce" target="_blank">Fun??o MULT</a> | FunctionResult | Multiplica seus argumentos |
| <a href="https://support.office.com/en-us/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94" target="_blank">Fun??o PRI.MAI?SCULA</a> | FunctionResult | Coloca a primeira letra de cada palavra em mai?scula em um valor de texto |
| <a href="https://support.office.com/en-us/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd" target="_blank">Fun??o VP</a> | FunctionResult | Retorna o valor presente de um investimento |
| <a href="https://support.office.com/en-us/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad" target="_blank">Fun??o QUARTIL.EXC</a> | FunctionResult | Retorna o quartil do conjunto de dados, com base em valores de percentil de 0..1, exclusive |
| <a href="https://support.office.com/en-us/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d" target="_blank">Fun??o QUARTIL.INC</a> | FunctionResult | Retorna o quartil de um conjunto de dados |
| <a href="https://support.office.com/en-us/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee" target="_blank">Fun??o QUOCIENTE</a> | FunctionResult | Retorna a parte inteira de uma divis?o |
| <a href="https://support.office.com/en-us/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf" target="_blank">Fun??o RADIANOS</a> | FunctionResult | Converte graus em radianos |
| <a href="https://support.office.com/en-us/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73" target="_blank">Fun??o ALEAT?RIO</a> | FunctionResult | Retorna um n?mero aleat?rio entre 0 e 1 |
| <a href="https://support.office.com/en-us/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685" target="_blank">Fun??o ALEAT?RIOENTRE</a> | FunctionResult | Retorna um n?mero aleat?rio entre os n?meros especificados |
| <a href="https://support.office.com/en-us/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a" target="_blank">Fun??o ORDEM.M?D</a> | FunctionResult | Retorna a posi??o de um n?mero em uma lista de n?meros |
| <a href="https://support.office.com/en-us/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40" target="_blank">Fun??o ORDEM.EQ</a> | FunctionResult | Retorna a posi??o de um n?mero em uma lista de n?meros |
| <a href="https://support.office.com/en-us/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce" target="_blank">Fun??o TAXA</a> | FunctionResult | Retorna a taxa de juros por per?odo de uma anuidade |
| <a href="https://support.office.com/en-us/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5" target="_blank">Fun??o RECEBIDO</a> | FunctionResult | Retorna a quantia recebida no vencimento de um t?tulo totalmente investido |
| <a href="https://support.office.com/en-us/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a" target="_blank">Fun??es MUDAR, SUBSTITUIRB</a> | FunctionResult | Muda os caracteres dentro do texto |
| <a href="https://support.office.com/en-us/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061" target="_blank">Fun??o REPT</a> | FunctionResult | Repete o texto um determinado n?mero de vezes |
| <a href="https://support.office.com/en-us/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f" target="_blank">Fun??es DIREITA, DIREITAB</a> | FunctionResult | Retorna os caracteres mais ? direita de um valor de texto |
| <a href="https://support.office.com/en-us/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5" target="_blank">Fun??o ROMANO</a> | FunctionResult | Converte um algarismo ar?bico em romano, como texto |
| <a href="https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c" target="_blank">Fun??o ARRED</a> | FunctionResult | Arredonda um n?mero at? uma quantidade especificada de d?gitos |
| <a href="https://support.office.com/en-us/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53" target="_blank">Fun??o ARREDONDAR.PARA.BAIXO</a> | FunctionResult | Arredonda um n?mero para baixo at? zero |
| <a href="https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7" target="_blank">Fun??o ARREDONDAR.PARA.CIMA</a> | FunctionResult | Arredonda um n?mero para cima afastando-o de zero |
| <a href="https://support.office.com/en-us/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597" target="_blank">Fun??o LINS</a> | FunctionResult | Retorna o n?mero de linhas em uma refer?ncia |
| <a href="https://support.office.com/en-us/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4" target="_blank">Fun??o TAXAJURO</a> | FunctionResult | Retorna uma taxa de juros equivalente para o crescimento de um investimento |
| <a href="https://support.office.com/en-us/article/RTD-function-e0cc001a-56f0-470a-9b19-9455dc0eb593" target="_blank">Fun??o RTD</a> | FunctionResult | Recupera dados em tempo real de um programa compat?vel com a automa??o COM |
| <a href="https://support.office.com/en-us/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7" target="_blank">Fun??o SEC</a> | FunctionResult | Retorna a secante de um ?ngulo |
| <a href="https://support.office.com/en-us/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f" target="_blank">Fun??o SECH</a> | FunctionResult | Retorna a secante hiperb?lica de um ?ngulo |
| <a href="https://support.office.com/en-us/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1" target="_blank">Fun??o SEGUNDO</a> | FunctionResult | Converte um n?mero de s?rie em um segundo |
| <a href="https://support.office.com/en-us/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637" target="_blank">Fun??o SOMASEQU?NCIA</a> | FunctionResult | Retorna a soma de uma s?rie polinomial baseada na f?rmula |
| <a href="https://support.office.com/en-us/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24" target="_blank">Fun??o PLAN</a> | FunctionResult | Retorna o n?mero da planilha referenciada |
| <a href="https://support.office.com/en-us/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b" target="_blank">Fun??o PLANS</a> | FunctionResult | Retorna o n?mero de planilhas em uma refer?ncia |
| <a href="https://support.office.com/en-us/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8" target="_blank">Fun??o SINAL</a> | FunctionResult | Retorna o sinal de um n?mero |
| <a href="https://support.office.com/en-us/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602" target="_blank">Fun??o SEN</a> | FunctionResult | Retorna o seno do ?ngulo fornecido |
| <a href="https://support.office.com/en-us/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7" target="_blank">Fun??o SENH</a> | FunctionResult | Retorna o seno hiperb?lico de um n?mero |
| <a href="https://support.office.com/en-us/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa" target="_blank">Fun??o DISTOR??O</a> | FunctionResult | Retorna a distor??o de uma distribui??o |
| <a href="https://support.office.com/en-us/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb" target="_blank">Fun??o DISTOR??O.P</a> | FunctionResult | Retorna a inclina??o de uma distribui??o com base em um preenchimento: uma caracteriza??o do grau de assimetria de uma distribui??o em torno de sua m?dia |
| <a href="https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8" target="_blank">Fun??o DPD</a> | FunctionResult | Retorna a deprecia??o em linha reta de um ativo durante um per?odo |
| <a href="https://support.office.com/en-us/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07" target="_blank">Fun??o MENOR</a> | FunctionResult | Retorna o menor valor k-?simo em um conjunto de dados |
| <a href="https://support.office.com/en-us/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf" target="_blank">Fun??o RAIZ</a> | FunctionResult | Retorna uma raiz quadrada positiva |
| <a href="https://support.office.com/en-us/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4" target="_blank">Fun??o RAIZPI</a> | FunctionResult | Retorna a raiz quadrada de (n?mero * pi) |
| <a href="https://support.office.com/en-us/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775" target="_blank">Fun??o PADRONIZAR</a> | FunctionResult | Retorna um valor normalizado |
| <a href="https://support.office.com/en-us/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285" target="_blank">Fun??o DESVPAD.P</a> | FunctionResult | Calcula o desvio padr?o com base no preenchimento completo |
| <a href="https://support.office.com/en-us/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23" target="_blank">Fun??o DESVPAD.A</a> | FunctionResult | Estima o desvio padr?o com base em uma amostra |
| <a href="https://support.office.com/en-us/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d" target="_blank">Fun??o DESVPADA</a> | FunctionResult | Estima o desvio padr?o com base em uma amostra, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c" target="_blank">Fun??o DESVPADPA</a> | FunctionResult | Calcula o desvio padr?o com base no preenchimento completo, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332" target="_blank">Fun??o SUBSTITUIR</a> | FunctionResult | Substitui um novo texto por um texto antigo em uma cadeia de texto |
| <a href="https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939" target="_blank">Fun??o SUBTOTAL</a> | FunctionResult | Retorna um subtotal em uma lista ou banco de dados |
| <a href="https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89" target="_blank">Fun??o SOMA</a> | FunctionResult | Soma seus argumentos |
| <a href="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b" target="_blank">Fun??o SOMASE</a> | FunctionResult | Adiciona as c?lulas especificadas por um determinado crit?rio |
| <a href="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b" target="_blank">Fun??o SOMASES</a> | FunctionResult | Adiciona as c?lulas de um intervalo que atendam a v?rios crit?rios |
| <a href="https://support.office.com/en-us/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307" target="_blank">Fun??o SOMAQUAD</a> | FunctionResult | Retorna a soma dos quadrados dos argumentos |
| <a href="https://support.office.com/en-us/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27" target="_blank">Fun??o SDA</a> | FunctionResult | Retorna a deprecia??o dos d?gitos da soma dos anos de um ativo para um per?odo especificado |
| <a href="https://support.office.com/en-us/article/T-function-fb83aeec-45e7-4924-af95-53e073541228" target="_blank">Fun??o T</a> | FunctionResult | Converte os argumentos em texto |
| <a href="https://support.office.com/en-us/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2" target="_blank">Fun??o DIST.T</a> | FunctionResult | Retorna os Pontos Percentuais (probabilidade) para a distribui??o t de Student |
| <a href="https://support.office.com/en-us/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28" target="_blank">Fun??o T.DIST.2T</a> | FunctionResult | Retorna os Pontos Percentuais (probabilidade) para a distribui??o t de Student |
| <a href="https://support.office.com/en-us/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda" target="_blank">Fun??o DIST.T.CD</a> | FunctionResult | Retorna a distribui??o t de Student |
| <a href="https://support.office.com/en-us/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e" target="_blank">Fun??o INV.T</a> | FunctionResult | Retorna o valor t da distribui??o t de Student como uma fun??o da probabilidade e dos graus de liberdade |
| <a href="https://support.office.com/en-us/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17" target="_blank">Fun??o T.INV.2T</a> | FunctionResult | Retorna o inverso da distribui??o t de Student |
| <a href="https://support.office.com/en-us/article/TAN-function-08851a40-179f-4052-b789-d7f699447401" target="_blank">Fun??o TAN</a> | FunctionResult | Retorna a tangente de um n?mero |
| <a href="https://support.office.com/en-us/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c" target="_blank">Fun??o TANH</a> | FunctionResult | Retorna a tangente hiperb?lica de um n?mero |
| <a href="https://support.office.com/en-us/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c" target="_blank">Fun??o OTN</a> | FunctionResult | Retorna o rendimento de um t?tulo equivalente a uma obriga??o do Tesouro |
| <a href="https://support.office.com/en-us/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2" target="_blank">Fun??o OTNVALOR</a> | FunctionResult | Retorna o pre?o por R$ 100,00 do valor nominal de uma obriga??o do Tesouro |
| <a href="https://support.office.com/en-us/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba" target="_blank">Fun??o OTNLUCRO</a> | FunctionResult | Retorna o rendimento de uma obriga??o do Tesouro |
| <a href="https://support.office.com/en-us/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c" target="_blank">Fun??o TEXTO</a> | FunctionResult | Formata um n?mero e o converte em texto |
| <a href="https://support.office.com/en-us/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457" target="_blank">Fun??o TEMPO</a> | FunctionResult | Retorna o n?mero de s?rie de uma hora espec?fica |
| <a href="https://support.office.com/en-us/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645" target="_blank">Fun??o VALOR.TEMPO</a> | FunctionResult | Converte um hor?rio na forma de texto em um n?mero de s?rie |
| <a href="https://support.office.com/en-us/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9" target="_blank">Fun??o HOJE</a> | FunctionResult | Retorna o n?mero de s?rie da data de hoje |
| <a href="https://support.office.com/en-us/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9" target="_blank">Fun??o ARRUMAR</a> | FunctionResult | Remove espa?os do texto |
| <a href="https://support.office.com/en-us/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3" target="_blank">Fun??o M?DIA.INTERNA</a> | FunctionResult | Retorna a m?dia do interior de um conjunto de dados |
| <a href="https://support.office.com/en-us/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb" target="_blank">Fun??o VERDADEIRO</a> | FunctionResult | Retorna o valor l?gico `TRUE` |
| <a href="https://support.office.com/en-us/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721" target="_blank">Fun??o TRUNC</a> | FunctionResult | Trunca um n?mero em um inteiro |
| <a href="https://support.office.com/en-us/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899" target="_blank">Fun??o TIPO</a> | FunctionResult | Retorna um n?mero indicando o tipo de dados de um valor |
| <a href="https://support.office.com/en-us/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8" target="_blank">Fun??o CARACTUNICODE</a> | FunctionResult | Retorna o caractere Unicode referenciado por determinado valor num?rico |
| <a href="https://support.office.com/en-us/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f" target="_blank">Fun??o UNICODE</a> | FunctionResult | Retorna o n?mero (ponto de c?digo) que corresponde ao primeiro caractere do texto |
| <a href="https://support.office.com/en-us/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6" target="_blank">Fun??o MAI?SCULA</a> | FunctionResult | Converte texto em mai?sculas |
| <a href="https://support.office.com/en-us/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2" target="_blank">Fun??o VALOR</a> | FunctionResult | Converte um argumento de texto em um n?mero |
| <a href="https://support.office.com/en-us/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a" target="_blank">Fun??o VAR.P</a> | FunctionResult | Calcula a varia??o com base no preenchimento inteiro |
| <a href="https://support.office.com/en-us/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b" target="_blank">Fun??o VAR.A</a> | FunctionResult | Estima a varia??o com base em uma amostra |
| <a href="https://support.office.com/en-us/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07" target="_blank">Fun??o VARA</a> | FunctionResult | Estima a varia??o com base em uma amostra, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96" target="_blank">Fun??o VARPA</a> | FunctionResult | Calcula a varia??o com base no preenchimento total, incluindo n?meros, texto e valores l?gicos |
| <a href="https://support.office.com/en-us/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73" target="_blank">Fun??o BDV</a> | FunctionResult | Retorna a deprecia??o de um ativo para um per?odo especificado ou parcial usando um m?todo de balan?o declinante |
| <a href="https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1" target="_blank">Fun??o PROCV</a> | FunctionResult | Procura na primeira coluna de uma matriz e se move ao longo da linha para retornar o valor de uma c?lula |
| <a href="https://support.office.com/en-us/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a" target="_blank">Fun??o DIA.DA.SEMANA</a> | FunctionResult | Converte um n?mero de s?rie em um dia da semana |
| <a href="https://support.office.com/en-us/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340" target="_blank">Fun??o N?MSEMANA</a> | FunctionResult | Converte um n?mero de s?rie em um n?mero que representa onde a semana cai numericamente em um ano |
| <a href="https://support.office.com/en-us/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db" target="_blank">Fun??o DIST.WEIBULL</a> | FunctionResult | Retorna a distribui??o de Weibull |
| <a href="https://support.office.com/en-us/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33" target="_blank">Fun??o DIATRABALHO</a> | FunctionResult | Retorna o n?mero de s?rie da data antes ou depois de um n?mero espec?fico de dias ?teis |
| <a href="https://support.office.com/en-us/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d" target="_blank">Fun??o DIATRABALHO.INTL</a> | FunctionResult | Retorna o n?mero de s?rie da data antes ou depois de um n?mero espec?fico de dias ?teis usando par?metros para indicar quais e quantos dias s?o de fim de semana |
| <a href="https://support.office.com/en-us/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d" target="_blank">Fun??o XIRR</a> | FunctionResult | Fornece a taxa interna de retorno para um programa de fluxos de caixa que n?o ? necessariamente peri?dico |
| <a href="https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7" target="_blank">Fun??o XVPL</a> | FunctionResult | Retorna o valor presente l?quido de um programa de fluxos de caixa que n?o ? necessariamente peri?dico |
| <a href="https://support.office.com/en-us/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37" target="_blank">Fun??o XOR</a> | FunctionResult | Retorna um OU exclusivo l?gico de todos os argumentos |
| <a href="https://support.office.com/en-us/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9" target="_blank">Fun??o ANO</a> | FunctionResult | Converte um n?mero de s?rie em um ano |
| <a href="https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8" target="_blank">Fun??o FRA??OANO</a> | FunctionResult | Retorna a fra??o do ano que representa o n?mero de dias entre a data_inicial e a data_final |
| <a href="https://support.office.com/en-us/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe" target="_blank">Fun??o LUCRO</a> | FunctionResult | Retorna o lucro de um t?tulo que paga juros peri?dicos |
| <a href="https://support.office.com/en-us/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7" target="_blank">Fun??o LUCRODESC</a> | FunctionResult | Retorna o rendimento anual de um t?tulo descontado. Por exemplo, uma obriga??o do Tesouro |
| <a href="https://support.office.com/en-us/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f" target="_blank">Fun??o LUCROVENC</a> | FunctionResult | Retorna o rendimento anual de um t?tulo que paga juros no vencimento |
| <a href="https://support.office.com/en-us/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee" target="_blank">Fun??o TESTE.Z</a> | FunctionResult | Retorna o valor de probabilidade unicaudal de um teste-z |

## <a name="see-also"></a>Confira tamb?m

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Especifica??o para abrir API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objeto de fun??es de planilha (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/functions)
