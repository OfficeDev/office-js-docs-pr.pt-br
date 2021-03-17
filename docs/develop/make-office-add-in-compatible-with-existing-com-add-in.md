---
title: Faça seu suplemento do Office ser compatível com um suplemento COM existente
description: Habilita a compatibilidade entre o seu Add-in do Office e o seu complemento COM equivalente.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836848"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Faça seu suplemento do Office ser compatível com um suplemento COM existente

Se você tiver um add-in COM existente, poderá criar funcionalidade equivalente em seu Add-in do Office, permitindo assim que sua solução seja executado em outras plataformas, como o Office na Web ou mac. Em alguns casos, seu Add-in do Office pode não ser capaz de fornecer toda a funcionalidade disponível no complemento COM correspondente. Nessas situações, o seu complemento COM pode oferecer uma experiência de usuário melhor no Windows do que o correspondente do Office Add-in pode fornecer.

Você pode configurar seu Add-in do Office para que, quando o complemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows executa o add-in COM em vez do Office Add-in. O complemento COM é chamado de "equivalente" porque o Office fará a transição perfeita entre o complemento COM e o Complemento do Office de acordo com o qual está instalado o computador de um usuário.

> [!NOTE]
> Esse recurso é suportado pelas seguintes plataformas, quando conectado a uma assinatura do Microsoft 365.
>
> - Excel, Word e PowerPoint na Web
> - Excel, Word e PowerPoint no Windows (versão 1904 ou posterior)
> - Excel, Word e PowerPoint no Mac (versão 13.329 ou posterior)
> - Outlook no Windows (versão 2102 ou posterior)

## <a name="specify-an-equivalent-com-add-in"></a>Especificar um complemento COM equivalente

### <a name="manifest"></a>Manifesto

> [!IMPORTANT]
> Aplica-se ao Excel, PowerPoint e Word. Suporte do Outlook em breve.

Para habilitar a compatibilidade entre o seu add-in do Office e o seu complemento COM, identifique o complemento COM equivalente no [manifesto](add-in-manifests.md) do seu Add-in do Office. Em seguida, o Office no Windows usará o complemento COM em vez do Office Add-in, se ambos estão instalados.

O exemplo a seguir mostra a parte do manifesto que especifica um complemento COM como um complemento equivalente. O valor do elemento identifica o complemento COM e o `ProgId` [elemento EquivalentAddins](../reference/manifest/equivalentaddins.md) deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento.

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
> Para obter informações sobre o complemento COM e a compatibilidade de UDF XLL, consulte Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário [XLL.](../excel/make-custom-functions-compatible-with-xll-udf.md)

### <a name="group-policy"></a>Política de grupo

> [!IMPORTANT]
> Aplica-se somente ao Outlook.

Para declarar compatibilidade entre o seu **add-in** da Web do Outlook e o complemento COM/VSTO, identifique o complemento COM equivalente na política de grupo Desative os complementos da Web do Outlook cujo complemento COM ou VSTO equivalente é instalado configurando no computador do usuário. Em seguida, o Outlook no Windows usará o complemento COM em vez do complemento da Web, se ambos estão instalados.

1. Baixe a ferramenta [Modelos Administrativos mais](https://www.microsoft.com/download/details.aspx?id=49030)recentes, preste atenção às Instruções de **Instalação da ferramenta.**
1. Abra o Editor de Política de Grupo Local (**gpedit.msc**).
1. Navegue **até Configuração do** Usuário Modelos  >     >  **Administrativos do Microsoft Outlook 2016**  >  **Diversos**.
1. Selecione a **configuração Desativar os complementos da Web do Outlook cujos complementos COM ou VSTO equivalentes estão instalados**.
1. Abra o link para editar a configuração de política.
1. Na caixa de diálogo **Os complementos da Web do Outlook para desativar**:
    1. Definir **o nome** do valor como o encontrado no manifesto do complemento da `Id` Web. **Importante**: *Não adicione* chaves ao redor da `{}` entrada.
    1. Definir **Valor** como `ProgId` o do complemento COM/VSTO equivalente.
    1. Selecione **OK** para colocar a atualização em vigor.
    ![Captura de tela mostrando a caixa de diálogo "Os complementos da Web do Outlook para desativar"](../images/outlook-deactivate-gpo-dialog.png)

## <a name="equivalent-behavior-for-users"></a>Comportamento equivalente para usuários

Quando um [complemento COM](#specify-an-equivalent-com-add-in)equivalente é especificado, o Office no Windows não exibirá a interface de usuário do seu Complemento do Office (UI) se o complemento COM equivalente estiver instalado. O Office oculta apenas os botões de faixa de opções do Office Add-in e não impede a instalação. Portanto, seu Complemento do Office ainda aparecerá nos seguintes locais na interface do usuário:

- Em **Meus complementos**
- Como entrada no gerenciador de faixa de opções (somente Excel, Word e PowerPoint)

> [!NOTE]
> A especificação de um complemento COM equivalente no manifesto não tem efeito em outras plataformas, como o Office na Web ou no Mac.

Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o Office Add-in.

### <a name="appsource-acquisition-of-an-office-add-in"></a>Aquisição do AppSource de um Add-in do Office

Se um usuário adquirir o Office Add-in do AppSource e o complemento COM equivalente já estiver instalado, o Office:

1. Instale o Office Add-in.
2. Ocultar a interface do usuário do Complemento do Office na faixa de opções.
3. Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.

### <a name="centralized-deployment-of-office-add-in"></a>Implantação centralizada do Office Add-in

Se um administrador implantar o Add-in do Office em seu locatário usando a implantação centralizada e o complemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver as alterações. Depois que o Office reiniciar, ele irá:

1. Instale o Office Add-in.
2. Ocultar a interface do usuário do Complemento do Office na faixa de opções.
3. Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Documento compartilhado com o Add-in incorporado do Office

Se um usuário tiver o add-in COM instalado e, em seguida, receber um documento compartilhado com o Complemento do Office incorporado, quando ele abrir o documento, o Office:

1. Solicitar que o usuário confie no Office Add-in.
2. Se for confiável, o Office Add-in será instalado.
3. Ocultar a interface do usuário do Complemento do Office na faixa de opções.

## <a name="other-com-add-in-behavior"></a>Outro comportamento de complemento COM

### <a name="excel-powerpoint-word"></a>Excel, PowerPoint, Word

Se um usuário desinstalar o complemento COM equivalente, o Office no Windows restaurará a interface do usuário do Office Add-in.

Depois de especificar um complemento COM equivalente para o seu Complemento do Office, o Office interrompe o processamento de atualizações para o seu Add-in do Office. Para adquirir as atualizações mais recentes do Office Add-in, o usuário deve primeiro desinstalar o complemento COM.

### <a name="outlook"></a>Outlook

O complemento COM/VSTO deve ser conectado quando o Outlook for iniciado para que o complemento da Web correspondente seja desabilitado.

Se o complemento COM/VSTO for desconectado durante uma sessão subsequente do Outlook, o complemento da Web provavelmente permanecerá desabilitado até que o Outlook seja reiniciado.

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md)
