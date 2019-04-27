---
title: Tornar o suplemento do Office compatível com um suplemento de COM existente
description: Habilitar a compatibilidade com um suplemento COM equivalente que tenha a mesma funcionalidade do suplemento do Office
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356831"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Tornar o suplemento do Office compatível com um suplemento de COM existente

Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Office para estender seus recursos de solução para outras plataformas, como online ou macOS. No enTanto, os suplementos do Office não possuem todas as funcionalidades disponíveis em suplementos de COM. O suplemento de COM pode fornecer uma experiência melhor do que o suplemento do Office no Windows no Excel, Word e PowerPoint.

Você pode configurar seu suplemento do Office para que, quando um suplemento COM equivalente já estiver instalado no computador do usuário, o Office execute o suplemento COM em vez do suplemento do Office. O suplemento de COM é chamado de "equivalente", pois o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Office, dependendo do que está instalado no Windows.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>Especificar um suplemento COM equivalente no manifesto

Para habilitar a compatibilidade com um suplemento de COM existente, identifique o suplemento COM equivalente no manifesto do suplemento do Office. O Office usará o suplemento COM em vez do suplemento do Office ao ser executado no Windows.

Especifique o `ProgID` do suplemento com equivalente. O Office usará a interface de usuário do suplemento COM em vez da interface do usuário do suplemento do Office quando o suplemento de COM estiver instalado.

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

Quando um suplemento COM equivalente é especificado no manifesto do suplemento do Office, o Office suprime a interface do usuário do suplemento do Office no Windows quando o suplemento COM equivalente está instalado. Isso não afeta a interface do usuário do suplemento do Office em outras plataformas como online ou macOS. O Office só oculta os botões da faixa de opções e não impede a instalação. Portanto, o suplemento do Office ainda aparecerá nos seguintes locais de interface do usuário:

- Em **meus suplementos** , pois ele é tecnicamente instalado.
- Como uma entrada no Gerenciador de faixa de opções.

Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Office.

### <a name="appsource-acquisition-of-an-office-add-in"></a>Aquisição do AppSource de um suplemento do Office

Se um usuário baixar o suplemento do Office do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:

1. Instalar o suplemento do Office.
2. Ocultar a interface do usuário do suplemento do Office na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="centralized-deployment-of-office-add-in"></a>Implantação centralizada do suplemento do Office

Se um administrador implantar o suplemento do Office em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário precisará reiniciar o Office para que ele possa ver as alterações. Após a reinicialização do Office, ela irá:

1. Instalar o suplemento do Office.
2. Ocultar a interface do usuário do suplemento do Office na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Documento compartilhado com o suplemento incorporado do Office

Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Office incorporado, quando abrir o documento, o Office irá:

1. Solicitar que o usuário confie no suplemento do Office.
2. Se confiável, o suplemento do Office será instalado.
3. Ocultar a interface do usuário do suplemento do Office na faixa de opções.

## <a name="other-com-add-in-behavior"></a>Outro comportamento de suplemento de COM

Se um usuário desinstala o suplemento de COM, o Office restaura a interface do usuário do suplemento do Office no Windows para o suplemento do Office instalado equivalente.

Após especificar um suplemento COM equivalente para o suplemento do Office, o Office interromperá o processamento de atualizações para seu suplemento do Office. O usuário deve desinstalar o suplemento de COM para obter as atualizações mais recentes para o suplemento do Office.

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md)
