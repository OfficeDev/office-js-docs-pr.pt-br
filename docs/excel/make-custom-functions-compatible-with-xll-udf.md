---
title: Estender funções personalizadas com funções definidas pelo usuário XLL
description: Habilitar a compatibilidade Excel funções definidas pelo usuário XLL que tenham funcionalidade equivalente às suas funções personalizadas
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 82d1120e68a69bee74a6fe1911bbd8d3ccb3fb00
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990709"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Estender funções personalizadas com funções definidas pelo usuário XLL

> [!NOTE]
> Um complemento XLL é um arquivo de Excel com a extensão de arquivo **.xll**. Um arquivo XLL é um tipo de arquivo DLL (biblioteca de links dinâmicos) que só pode ser aberto por Excel. Os arquivos de complemento XLL devem ser gravados em C ou C++. Consulte [Desenvolvendo Excel XLLs](/office/client-developer/excel/developing-excel-xlls) para saber mais.

Se você tiver os Excel XLL existentes, poderá criar complementos de função personalizada equivalentes usando Excel API JavaScript do Excel para estender seus recursos de solução para outras plataformas, como Excel na Web ou em um Mac. No entanto, Excel de API JavaScript não têm todas as funcionalidades disponíveis em complementos XLL. Dependendo da funcionalidade que sua solução usa, o complemento XLL pode oferecer uma experiência melhor do que o Excel de API JavaScript do Excel no Windows.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>Especificar XLL equivalente no manifesto

Para habilitar a compatibilidade com um complemento XLL existente, identifique o complemento XLL equivalente no manifesto do seu Excel de API JavaScript. Excel usarão as funções do complemento XLL em vez de funções personalizadas do seu Excel de api JavaScript ao ser executado no Windows.

Para definir o complemento XLL equivalente para suas funções personalizadas, especifique o `FileName` arquivo XLL. Quando o usuário abre uma pasta de trabalho com funções do arquivo XLL, Excel converte as funções em funções compatíveis. A pasta de trabalho usa o arquivo XLL quando aberta no Excel no Windows e usará funções personalizadas do seu complemento da API JavaScript do Excel quando aberta na Web ou em um Mac.

O exemplo a seguir mostra como especificar um add-in COM e um complemento XLL como equivalentes em um arquivo de manifesto de manifesto do Excel API JavaScript. Muitas vezes, você especificará ambos. Para completar, este exemplo mostra ambos no contexto. Eles são identificados por `ProgId` seus e `FileName` respectivamente. O `EquivalentAddins` elemento deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento. Para obter mais informações sobre compatibilidade com o complemento COM, consulte [Make your Office Add-in compatible with an existing COM add-in](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> Se um Excel de API JavaScript declarar suas funções personalizadas como compatíveis com um complemento XLL, alterar o manifesto posteriormente poderá quebrar a pasta de trabalho de um usuário porque alterará o formato do arquivo.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportamento de função personalizado para funções compatíveis com XLL

As funções XLL de um complemento são convertidas em funções personalizadas compatíveis com XLL quando uma planilha é aberta e há um complemento equivalente disponível. Na próxima salvação, as funções XLL são escritas no arquivo em um modo compatível para que funcionem com o complemento XLL e com funções personalizadas do complemento da API JavaScript do Excel do Excel (quando em outras plataformas).

A tabela a seguir compara os recursos entre funções definidas pelo usuário XLL, funções personalizadas compatíveis com XLL e Excel funções personalizadas do complemento da API JavaScript.

|         |Função definida pelo usuário XLL |Funções personalizadas compatíveis com XLL |Excel Função personalizada do add-in da API JavaScript |
|---------|---------|---------|---------|
| **Plataformas com suporte** | Windows | Windows, macOS, navegador da Web | Windows, macOS, navegador da Web |
| **Formatos de arquivo com suporte** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Preenchimento automático de fórmula** | Não | Sim | Sim |
| **Streaming** | Possível por meio de retorno de chamada xlfRTD e XLL. | Sim | Sim |
| **Localização de funções** | Não | Não. O Nome e a ID devem corresponder às funções existentes da XLL. | Sim |
| **Funções voláteis** | Sim | Sim | Sim |
| **Suporte a recálculo com vários threads** | Sim | Sim | Sim |
| **Comportamento de cálculo** | Sem interface do usuário. Excel pode ser não responsivo durante o cálculo. | Os usuários verão #BUSY! até que um resultado seja retornado. | Os usuários verão #BUSY! até que um resultado seja retornado. |
| **Conjuntos de requisitos** | N/D | CustomFunctions 1.1 e posterior | CustomFunctions 1.1 e posterior |

## <a name="see-also"></a>Ver também

- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
