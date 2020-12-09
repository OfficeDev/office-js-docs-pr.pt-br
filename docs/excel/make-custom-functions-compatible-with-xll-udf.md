---
title: Estender funções personalizadas com funções definidas pelo usuário XLL
description: Habilitar a compatibilidade com as funções definidas pelo usuário do Excel XLL que possuem funcionalidade equivalente às suas funções personalizadas
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: c34dcf5ef546fa0f337b2cbd11cca7d5e25e2de3
ms.sourcegitcommit: fecad2afa7938d7178456c11ba52b558224813b4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49603775"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>Estender funções personalizadas com funções definidas pelo usuário XLL

Se você tiver os XLLs do Excel existentes, poderá criar funções personalizadas equivalentes em um suplemento do Excel para estender seus recursos de solução para outras plataformas, como online ou em um Mac. No entanto, os suplementos do Excel não possuem todas as funcionalidades disponíveis em XLLs. Dependendo da funcionalidade que sua solução usa, o XLL pode fornecer uma experiência melhor do que as funções personalizadas do suplemento do Excel no Excel no Windows.

> [!NOTE]
> O suplemento de COM e a compatibilidade do XLL UDF são compatíveis com as seguintes plataformas, quando conectados a uma assinatura do Microsoft 365:
> - Excel Online
> - Excel no Windows (versão 1904 ou posterior)
> - Excel no Mac (versão 13,329 ou posterior)
>
> Para usar o suplemento de COM e a compatibilidade do XLL UDF no Excel na Web, faça logon usando sua assinatura do Microsoft 365 ou uma [conta da Microsoft](https://account.microsoft.com/account). Se você ainda não tem uma assinatura do Microsoft 365, é possível uma assinatura gratuita, de 90 dias, redimensionada para o Microsoft 365, participando do [programa de desenvolvedor do microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="specify-equivalent-xll-in-the-manifest"></a>Especificar o XLL equivalente no manifesto

Para habilitar a compatibilidade com um XLL existente, identifique o XLL equivalente no manifesto do suplemento do Excel. O Excel usará as funções do XLL em vez de suas funções personalizadas do suplemento do Excel ao executar o Windows.

Para definir o XLL equivalente para suas funções personalizadas, especifique o `FileName` do XLL. Quando o usuário abre uma pasta de trabalho com funções do XLL, o Excel converte as funções em funções compatíveis. Em seguida, a pasta de trabalho usa o XLL quando aberto no Excel no Windows, e ele usará as funções personalizadas do seu suplemento do Excel quando ele for aberto online ou em um Mac.

O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente. Em geral, você especifica ambos. Para concluir, este exemplo mostra tanto no contexto. Eles são identificados por seus `ProgId` e, `FileName` respectivamente. O `EquivalentAddins` elemento deve ser posicionado imediatamente antes da `VersionOverrides` marca de fechamento. Para obter mais informações sobre a compatibilidade do suplemento COM, consulte [tornar o suplemento do Excel compatível com um suplemento de com existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

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
> Se um suplemento declarar suas funções personalizadas para serem compatíveis com XLL, alterar o manifesto posteriormente poderá quebrar a pasta de trabalho de um usuário, pois ele alterará o formato de arquivo.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportamento de função personalizada para funções compatíveis com XLL

As funções XLL de um suplemento são convertidas em funções personalizadas compatíveis com o XLL quando uma planilha é aberta e há um suplemento equivalente disponível. Na próxima vez que você salvar, as funções XLL serão gravadas no arquivo em um modo compatível para que funcionem com as funções personalizadas XLL e suplemento do Excel (quando estiverem em outras plataformas).

A tabela a seguir compara os recursos nas funções de XLL definidas pelo usuário, funções personalizadas compatíveis e funções personalizadas de suplemento do Excel.

|         |Função de XLL definida pelo usuário |Funções personalizadas compatíveis com XLL |Função personalizada de suplemento do Excel |
|---------|---------|---------|---------|
| **Plataformas compatíveis** | Windows | Windows, macOS, navegador da Web | Windows, macOS, navegador da Web |
| **Formatos de arquivo com suporte** | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| **Preenchimento automático de fórmula** | Não | Sim | Sim |
| **Streaming** | Possível via xlfRTD e o retorno de chamada XLL. | Sim | Sim |
| **Localização de funções** | Não | Não. O nome e a ID devem corresponder às funções de XLL existentes. | Sim |
| **Funções voláteis** | Sim | Sim | Sim |
| **Suporte para recálculo de vários encadeamentos** | Sim | Sim | Sim |
| **Comportamento de cálculo** | Nenhuma interface do usuário. O Excel pode não responder durante o cálculo. | Os usuários verão #BUSY! até que um resultado seja retornado. | Os usuários verão #BUSY! até que um resultado seja retornado. |
| **Conjuntos de requisitos** | N/D | CustomFunctions 1,1 e posterior | CustomFunctions 1,1 e posterior |

## <a name="see-also"></a>Confira também

- [Tornar seu suplemento do Excel compatível com um suplemento de COM existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
