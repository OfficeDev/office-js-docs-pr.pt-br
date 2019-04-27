---
title: Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL
description: Habilitar a compatibilidade com as funções definidas pelo usuário do Excel XLL que possuem funcionalidade equivalente às suas funções personalizadas
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356833"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a>Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL

Se você tiver os XLLs do Excel existentes, poderá criar funções personalizadas equivalentes em um suplemento do Office para estender seus recursos de solução para outras plataformas, como online ou macOS. No enTanto, os suplementos do Office não possuem todas as funcionalidades disponíveis em XLLs. Dependendo da funcionalidade que sua solução usa, o XLL pode fornecer uma experiência melhor do que as funções personalizadas do suplemento do Office no Excel para Windows.

Você pode configurar seu suplemento do Office para que, quando um XLL equivalente já estiver instalado no computador do usuário, o Excel execute o XLL, em vez de suas funções personalizadas do suplemento do Office. O XLL é chamado de equivalente porque o Excel faz a transição contínua entre as funções de personalização XLL e suplemento do Office, dependendo da que está instalada no Windows.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>Especificar o XLL equivalente no manifesto

Para habilitar a compatibilidade com um XLL existente, identifique o XLL equivalente no manifesto do suplemento do Office. Em seguida, o Excel usará as funções do XLL em vez de suas funções personalizadas do suplemento do Office durante a execução no Windows.

Para definir o XLL equivalente para suas funções personalizadas, especifique o `FileName` do XLL. Quando o usuário abre uma pasta de trabalho com funções do XLL, o Excel converte as funções em funções compatíveis. Em seguida, a pasta de trabalho usa o XLL quando aberto no Excel no Windows, e ele usará as funções personalizadas do seu suplemento do Office quando for aberto online ou no macOS.

O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente. Em geral, você especifica tanto tanto quanto à integridade este exemplo mostra tanto no contexto. Eles são identificados por seus `ProgID` e `FileName` , respectivamente. Para obter mais informações sobre a compatibilidade do suplemento COM, confira [tornar o suplemento do Office compatível com um suplemento de com existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md).

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> Se um suplemento declarar suas funções personalizadas para serem compatíveis com XLL, alterar o manifesto posteriormente poderá quebrar a pasta de trabalho de um usuário, pois ele alterará o formato de arquivo.

## <a name="office-add-in-updates"></a>Atualizações de suplementos do Office

Depois de especificar um XLL equivalente para o suplemento do Office, o Excel interrompe o processamento de atualizações para seu suplemento do Office. O usuário deve desinstalar o XLL para obter as atualizações mais recentes para o suplemento do Office.

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>Comportamento de função personalizada para funções compatíveis com XLL

Quando uma planilha é aberta contendo funções XLL para as quais há também um suplemento equivalente, as funções do XLL são convertidas em funções personalizadas compatíveis com XLL. No próximo salvamento, eles são gravados no arquivo em um modo compatível para que funcionem com as funções personalizadas do XLL e do suplemento do Office (quando em outras plataformas).

A tabela a seguir compara os recursos nas funções de XLL definidas pelo usuário, funções personalizadas compatíveis e funções personalizadas do suplemento do Office.

|         |Função de XLL definida pelo usuário |Funções personalizadas compatíveis com XLL |Função personalizada de suplemento do Office |
|---------|---------|---------|---------|
| Plataformas com suporte | Windows | Windows, macOS, Excel online | Windows, macOS, Excel online |
| Formatos de arquivo com suporte | XLSX, XLSB, XLSM, XLS | XLSX, XLSB, XLSM | XLSX, XLSB, XLSM |
| Preenchimento automático de fórmula | Não | Sim | Sim |
| Streaming | Possível via xlfRTD e o retorno de chamada XLL. | Sim | Sim |
| Localização de funções | Não | Não. O nome e a ID devem corresponder às funções de XLL existentes. | Sim |
| Funções voláteis | Sim | Sim | Sim |
| Suporte para recálculo de vários encadeamentos | Sim | Sim | Sim |
| Comportamento de cálculo | Nenhuma interface do usuário. O Excel pode não responder durante o cálculo. | Os usuários verão #BUSY! até que um resultado seja retornado. | Os usuários verão #BUSY! até que um resultado seja retornado. |
| Conjuntos de requisitos | N/D | CustomFunctions 1,1 somente | CustomFunctions 1,1 e posterior |

## <a name="see-also"></a>Confira também

- [Tornar o suplemento do Office compatível com um suplemento de COM existente](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
- [Log de alteração de funções personalizadas](custom-functions-changelog.md)
- [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)