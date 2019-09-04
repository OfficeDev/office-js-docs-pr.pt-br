---
title: Torne o seu suplemento do Office compatível com um suplemento COM existente
description: Habilitar a compatibilidade entre o suplemento do Office e o suplemento COM equivalente
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: bd419d059abd51f969affe107e8ec54e66bdac7f
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715610"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Torne o seu suplemento do Office compatível com um suplemento COM existente

Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Office, permitindo assim que sua solução seja executada em outras plataformas, como o Office na Web ou o Office no Mac. Em alguns casos, o suplemento do Office pode não ser capaz de fornecer toda a funcionalidade que está disponível no suplemento COM correspondente. Nessas situações, o suplemento COM pode fornecer uma experiência de usuário melhor no Windows do que o suplemento do Office correspondente pode fornecer.

Você pode configurar seu suplemento do Office para que, quando o suplemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows execute o suplemento COM em vez do suplemento do Office. O suplemento de COM é chamado de "equivalente" porque o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Office de acordo com o qual está instalado o computador de um usuário.

> [!NOTE]
> Este recurso é suportado pelas seguintes plataformas, quando conectado a uma assinatura do Office 365:
> - Excel, Word e PowerPoint na Web
> - Excel, Word e PowerPoint no Windows (versão 1904 ou posterior)
> - Excel, Word e PowerPoint no Mac (versão 13,329 ou posterior)

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>Especificar um suplemento COM equivalente no manifesto

Para habilitar a compatibilidade entre o suplemento do Office e o suplemento de COM, identifique o suplemento COM equivalente no [manifesto](add-in-manifests.md) do suplemento do Office. O Office no Windows usará o suplemento COM em vez do suplemento do Office, se eles estiverem instalados.

O exemplo a seguir mostra a parte do manifesto que especifica um suplemento de COM como um suplemento equivalente. O valor do `ProgId` elemento identifica o suplemento de com e o `EquivalentAddins` elemento deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento.

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> Para obter informações sobre o suplemento de COM e a compatibilidade do XLL UDF, confira [tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário do XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).

## <a name="equivalent-behavior-for-users"></a>Comportamento equivalente para usuários

Quando um suplemento COM equivalente é especificado no manifesto do suplemento do Office, o Office no Windows não exibirá a interface do usuário do suplemento do Office se o suplemento COM equivalente estiver instalado. O Office só oculta os botões da faixa de opções do suplemento do Office e não impede a instalação. Portanto, o suplemento do Office ainda aparecerá nos seguintes locais dentro da interface do usuário:

- Em **meus suplementos**
- Como uma entrada no Gerenciador de faixa de opções

> [!NOTE]
> A especificação de um suplemento COM equivalente no manifesto não tem efeito sobre outras plataformas como o Office na Web ou Mac.

Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Office.

### <a name="appsource-acquisition-of-an-office-add-in"></a>Aquisição do AppSource de um suplemento do Office

Se um usuário adquire o suplemento do Office do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:

1. Instalar o suplemento do Office.
2. Ocultar a interface do usuário do suplemento do Office na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="centralized-deployment-of-office-add-in"></a>Implantação centralizada do suplemento do Office

Se um administrador implantar o suplemento do Office em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver as alterações. Após a reinicialização do Office, ela irá:

1. Instalar o suplemento do Office.
2. Ocultar a interface do usuário do suplemento do Office na faixa de opções.
3. Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Documento compartilhado com o suplemento incorporado do Office

Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Office incorporado, quando abrir o documento, o Office irá:

1. Solicitar que o usuário confie no suplemento do Office.
2. Se confiável, o suplemento do Office será instalado.
3. Ocultar a interface do usuário do suplemento do Office na faixa de opções.

## <a name="other-com-add-in-behavior"></a>Outro comportamento de suplemento de COM

Se um usuário desinstalar o suplemento COM equivalente, o Office no Windows restaura a interface do usuário do suplemento do Office.

Depois de especificar um suplemento COM equivalente para o suplemento do Office, o Office interrompe o processamento de atualizações para seu suplemento do Office. Para adquirir as atualizações mais recentes para o suplemento do Office, o usuário deve primeiro desinstalar o suplemento de COM.

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md)
