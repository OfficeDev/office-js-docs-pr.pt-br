---
title: Estender funções personalizadas com funções definidas pelo usuário XLL
description: Habilitar a compatibilidade com funções definidas pelo usuário do Excel XLL que tenham funcionalidade equivalente às suas funções personalizadas
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 32146e7eebb963e8d800b619ef052457e40f2ac6
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836813"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Estender funções personalizadas com funções definidas pelo usuário XLL

Se você tiver XLLs do Excel existentes, poderá criar funções personalizadas equivalentes em um complemento do Excel para estender seus recursos de solução para outras plataformas, como online ou em um Mac. No entanto, os complementos do Excel não têm toda a funcionalidade disponível em XLLs. Dependendo da funcionalidade que sua solução usa, a XLL pode oferecer uma experiência melhor do que as funções personalizadas do complemento do Excel no Excel no Windows.

> [!NOTE]
> O complemento COM e a compatibilidade UDF XLL são compatíveis com as seguintes plataformas, quando conectadas a uma assinatura do Microsoft 365:
>
> - Excel Online
> - Excel no Windows (versão 1904 ou posterior)
> - Excel no Mac (versão 13.329 ou posterior)
>
> Para usar o complemento COM e a compatibilidade UDF XLL no Excel na Web, faça logon usando sua assinatura do Microsoft 365 ou uma [conta da Microsoft.](https://account.microsoft.com/account) Se você ainda não tiver uma assinatura do Microsoft 365, poderá ter uma assinatura do Microsoft [365](https://developer.microsoft.com/office/dev-program)renovável de 90 dias gratuita, juntando-se ao programa de desenvolvedor do Microsoft 365 .

## <a name="specify-equivalent-xll-in-the-manifest"></a>Especificar XLL equivalente no manifesto

Para habilitar a compatibilidade com uma XLL existente, identifique a XLL equivalente no manifesto do seu complemento do Excel. O Excel usará as funções XLL em vez de funções personalizadas do seu complemento do Excel ao ser executado no Windows.

Para definir a XLL equivalente para suas funções personalizadas, especifique o `FileName` da XLL. Quando o usuário abre uma planilha com funções da XLL, o Excel converte as funções em funções compatíveis. A planilha usa a XLL quando aberta no Excel no Windows e usará funções personalizadas do seu complemento do Excel quando aberta online ou em um Mac.

O exemplo a seguir mostra como especificar um complemento COM e uma XLL como equivalente. Muitas vezes, você especificará ambos. Para completar, este exemplo mostra ambos no contexto. Eles são identificados por `ProgId` seus e `FileName` respectivamente. O `EquivalentAddins` elemento deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento. Para obter mais informações sobre compatibilidade com o complemento COM, consulte Tornar seu Add-in do Office compatível com um [add-in COM existente.](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)

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
> Se um complemento declarar que suas funções personalizadas são compatíveis com XLL, alterar o manifesto posteriormente pode quebrar a pasta de trabalho do usuário porque ele alterará o formato de arquivo.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportamento de função personalizado para funções compatíveis com XLL

As funções XLL de um complemento são convertidas em funções personalizadas compatíveis com XLL quando uma planilha é aberta e há um complemento equivalente disponível. Na próxima salvação, as funções XLL são escritas no arquivo em um modo compatível para que elas funcionem com as funções personalizadas do complemento XLL e do Excel (quando em outras plataformas).

A tabela a seguir compara recursos entre funções definidas pelo usuário XLL, funções personalizadas compatíveis com XLL e funções personalizadas de complemento do Excel.

|         |Função definida pelo usuário XLL |Funções personalizadas compatíveis com XLL |Função personalizada do complemento do Excel |
|---------|---------|---------|---------|
| **Plataformas compatíveis** | Windows | Windows, macOS, navegador da Web | Windows, macOS, navegador da Web |
| **Formatos de arquivo com suporte** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Preenchimento automático de fórmula** | Não | Sim | Sim |
| **Streaming** | Possível por meio de retorno de chamada xlfRTD e XLL. | Sim | Sim |
| **Localização de funções** | Não | Não. O Nome e a ID devem corresponder às funções existentes da XLL. | Sim |
| **Funções voláteis** | Sim | Sim | Sim |
| **Suporte a recálculo com vários threads** | Sim | Sim | Sim |
| **Comportamento de cálculo** | Sem interface do usuário. O Excel pode não responder durante o cálculo. | Os usuários verão #BUSY! até que um resultado seja retornado. | Os usuários verão #BUSY! até que um resultado seja retornado. |
| **Conjuntos de requisitos** | N/D | CustomFunctions 1.1 e posterior | CustomFunctions 1.1 e posterior |

## <a name="see-also"></a>Confira também

- [Torne o seu suplemento do Office compatível com um suplemento COM existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
