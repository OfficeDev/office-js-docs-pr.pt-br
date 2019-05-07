---
title: Tornar seu suplemento do Excel compatível com um suplemento de COM existente
description: Habilitar a compatibilidade com um suplemento COM equivalente que tenha a mesma funcionalidade do seu suplemento do Excel
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628169"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>Tornar o suplemento do Office compatível com um suplemento de COM existente (visualização)

Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Excel para estender seus recursos de solução para outras plataformas, como online ou macOS. No entanto, os suplementos do Excel não possuem todas as funcionalidades disponíveis em suplementos de COM. O suplemento de COM pode fornecer uma experiência melhor do que o suplemento do Excel no Windows.

Você pode configurar seu suplemento do Excel para que, quando um suplemento COM equivalente já estiver instalado no computador do usuário, o Office execute o suplemento COM em vez do suplemento do Excel. O suplemento de COM é chamado de "equivalente", pois o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Excel, dependendo do que está instalado no Windows.

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>Especificar um suplemento COM equivalente no manifesto

Para habilitar a compatibilidade com um suplemento de COM existente, identifique o suplemento COM equivalente no manifesto do suplemento do Excel. Em seguida, o Office usará o suplemento COM em vez do seu suplemento do Excel ao executar o Windows.

Especifique o `ProgID` do suplemento com equivalente. O Office usará a interface de usuário do suplemento COM em vez da interface do usuário do suplemento do Excel quando o suplemento de COM estiver instalado.

O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente. Em geral, você especifica tanto tanto quanto à integridade este exemplo mostra tanto no contexto. Eles são identificados por seus `ProgID` e `FileName` , respectivamente. Para obter mais informações sobre a compatibilidade XLL, consulte [tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário do XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).

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

## <a name="equivalent-behavior-for-users"></a>Comportamento equivalente para usuários

Quando um suplemento COM equivalente é especificado no manifesto de suplemento do Excel, o Office suprime sua interface do usuário do suplemento do Excel no Windows quando o suplemento COM equivalente está instalado. Isso não afeta a interface do usuário do seu suplemento do Excel em outras plataformas como online ou macOS. O Office só oculta os botões da faixa de opções e não impede a instalação. Portanto, o suplemento do Excel ainda aparecerá nos seguintes locais de interface do usuário:

- Em **meus suplementos** , pois ele é tecnicamente instalado.
- Como uma entrada no Gerenciador de faixa de opções.

Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Excel.

### <a name="appsource-acquisition-of-an-excel-add-in"></a>AppSource aquisição de um suplemento do Excel

Se um usuário baixar o suplemento do Excel do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:

1. Instalar o suplemento do Excel.
2. Ocultar a interface do usuário do suplemento do Excel na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="centralized-deployment-of-excel-add-in"></a>Implantação centralizada do suplemento do Excel

Se um administrador implantar o suplemento do Excel em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário precisará reiniciar o Office para que ele possa ver as alterações. Após a reinicialização do Office, ela irá:

1. Instalar o suplemento do Excel.
2. Ocultar a interface do usuário do suplemento do Excel na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="document-shared-with-embedded-excel-add-in"></a>Documento compartilhado com o suplemento incorporado do Excel

Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Excel incorporado, quando abrir o documento, o Office irá:

1. Solicitar que o usuário confie no suplemento do Excel.
2. Se confiável, o suplemento do Excel será instalado.
3. Ocultar a interface do usuário do suplemento do Excel na faixa de opções.

## <a name="other-com-add-in-behavior"></a>Outro comportamento de suplemento de COM

Se um usuário desinstala o suplemento de COM, o Office restaura a interface de usuário do suplemento do Excel no Windows para o suplemento do Excel instalado equivalente.

Depois de especificar um suplemento de COM equivalente para seu suplemento do Excel, o Office interrompe o processamento de atualizações para seu suplemento do Excel. O usuário deve desinstalar o suplemento de COM para obter as atualizações mais recentes para o suplemento do Excel.

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md)
