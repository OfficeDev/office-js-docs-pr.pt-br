---
title: Estender funções personalizadas com funções definidas pelo usuário XLL
description: Habilitar a compatibilidade com as funções definidas pelo usuário do Excel XLL que possuem funcionalidade equivalente às suas funções personalizadas (visualização)
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: fa80597d48859a29bcdbb5b900c251382dc87caa
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771650"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions-preview"></a>Estender funções personalizadas com funções definidas pelo usuário XLL (visualização)

Se você tiver os XLLs do Excel existentes, poderá criar funções personalizadas equivalentes em um suplemento do Excel para estender seus recursos de solução para outras plataformas, como online ou macOS. No entanto, os suplementos do Excel não possuem todas as funcionalidades disponíveis em XLLs. Dependendo da funcionalidade que sua solução usa, o XLL pode fornecer uma experiência melhor do que as funções personalizadas do suplemento do Excel no Excel no Windows.

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>Especificar o XLL equivalente no manifesto

Para habilitar a compatibilidade com um XLL existente, identifique o XLL equivalente no manifesto do suplemento do Excel. Em seguida, o Excel usará as funções do XLL em vez de suas funções personalizadas do suplemento do Excel ao executar o Windows.

Para definir o XLL equivalente para suas funções personalizadas, especifique o `FileName` do XLL. Quando o usuário abre uma pasta de trabalho com funções do XLL, o Excel converte as funções em funções compatíveis. Em seguida, a pasta de trabalho usa o XLL quando aberto no Excel no Windows, e ele usará as funções personalizadas do seu suplemento do Excel quando ele for aberto online ou no macOS.

O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente. Em geral, você especifica tanto tanto quanto à integridade este exemplo mostra tanto no contexto. Eles são identificados por seus `ProgId` e `FileName` , respectivamente. O `EquivalentAddins` elemento deve ser posicionado imediatamente antes da `VersionOverrides` marca de fechamento. Para obter mais informações sobre a compatibilidade do suplemento COM, consulte [tornar o suplemento do Excel compatível com um suplemento de com existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

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
  <EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> Se um suplemento declarar suas funções personalizadas para serem compatíveis com XLL, alterar o manifesto posteriormente poderá quebrar a pasta de trabalho de um usuário, pois ele alterará o formato de arquivo.

## <a name="excel-add-in-updates"></a>Atualizações de suplementos do Excel

Depois de especificar um XLL equivalente para seu suplemento do Excel, o Excel interrompe o processamento de atualizações para seu suplemento do Excel. O usuário deve desinstalar o XLL para obter as atualizações mais recentes para o suplemento do Excel.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportamento de função personalizada para funções compatíveis com XLL

Quando uma planilha é aberta contendo funções XLL para as quais há também um suplemento equivalente, as funções do XLL são convertidas em funções personalizadas compatíveis com XLL. Na próxima vez que você salvar, eles serão gravados no arquivo em um modo compatível para que eles funcionem com as funções personalizadas do XLL e do suplemento do Excel (quando em outras plataformas).

A tabela a seguir compara os recursos nas funções de XLL definidas pelo usuário, funções personalizadas compatíveis e funções personalizadas de suplemento do Excel.

|         |Função de XLL definida pelo usuário |Funções personalizadas compatíveis com XLL |Função personalizada de suplemento do Excel |
|---------|---------|---------|---------|
| Plataformas com suporte | Windows | Windows, macOS, Excel online | Windows, macOS, Excel online |
| Formatos de arquivo com suporte | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Preenchimento automático de fórmula | Não | Sim | Sim |
| Streaming | Possível via xlfRTD e o retorno de chamada XLL. | Não | Sim |
| Localização de funções | Não | Não. O nome e a ID devem corresponder às funções de XLL existentes. | Sim |
| Funções voláteis | Sim | Sim | Sim |
| Suporte para recálculo de vários encadeamentos | Sim | Sim | Sim |
| Comportamento de cálculo | Nenhuma interface do usuário. O Excel pode não responder durante o cálculo. | Os usuários verão #BUSY! até que um resultado seja retornado. | Os usuários verão #BUSY! até que um resultado seja retornado. |
| Conjuntos de requisitos | N/D | CustomFunctions 1,1 e posterior | CustomFunctions 1,1 e posterior |

## <a name="see-also"></a>Confira também

- [Tornar seu suplemento do Excel compatível com um suplemento de COM existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
