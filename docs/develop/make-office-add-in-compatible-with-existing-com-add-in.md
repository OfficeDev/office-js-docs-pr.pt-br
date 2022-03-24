---
title: Faça seu suplemento do Office ser compatível com um suplemento COM existente
description: Habilita a compatibilidade entre o seu Office e o seu complemento COM equivalente.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: f78f41532f916dc5df43cf5a6d4e455b6f16864f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743797"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>Faça seu suplemento do Office ser compatível com um suplemento COM existente

Se você tiver um complemento COM existente, poderá criar funcionalidade equivalente no seu Office Add-in, permitindo assim que sua solução seja executado em outras plataformas, como Office na Web ou Mac. Em alguns casos, seu Office de usuário pode não ser capaz de fornecer toda a funcionalidade disponível no complemento COM correspondente. Nessas situações, o seu complemento COM pode fornecer uma melhor experiência do usuário Windows do que o Office que o Add-in pode fornecer.

Você pode configurar seu Office Add-in para que, quando o complemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows executa o add-in COM em vez do Office Add-in. O add-in COM é chamado de "equivalente" porque o Office fará a transição perfeita entre o complemento COM e o Office De acordo com o qual está instalado o computador de um usuário.

[!INCLUDE [Support note for equivalent add-ins feature](../includes/equivalent-add-in-support-note.md)]

## <a name="specify-an-equivalent-com-add-in"></a>Especificar um complemento COM equivalente

### <a name="manifest"></a>Manifesto

> [!IMPORTANT]
> Aplica-se Excel, Outlook, PowerPoint e Word.

Para habilitar a compatibilidade entre o seu Office e o complemento COM, identifique o complemento COM equivalente no manifesto do seu Office Add-in[](add-in-manifests.md). Em seguida Office no Windows usará o add-in COM em vez do Office de Office, se ambos estão instalados.

O exemplo a seguir mostra a parte do manifesto que especifica um complemento COM como um complemento equivalente. O valor do elemento `ProgId` identifica o complemento COM e o [elemento EquivalentAddins](../reference/manifest/equivalentaddins.md) deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento.

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
> Para obter informações sobre o complemento COM e a compatibilidade de UDF XLL, consulte Tornar suas funções personalizadas compatíveis com [funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md). Não aplicável para Outlook.

### <a name="group-policy"></a>Política de grupo

> [!IMPORTANT]
> Aplica-se Outlook somente.

Para declarar compatibilidade entre o seu add-in da Web do Outlook e o complemento COM/VSTO, identifique o complemento COM equivalente na política de grupo **Desative os complementos da Web do Outlook cujos complementos equivalentes com ou VSTO estão instalados** configurando-se no computador do usuário. Em seguida, Outlook no Windows usará o add-in COM em vez do complemento da Web, se ambos estão instalados.

1. Baixe a ferramenta [Modelos Administrativos mais recentes](https://www.microsoft.com/download/details.aspx?id=49030), atenta às Instruções de **Instalação da ferramenta**.
1. Abra o Editor de Política de Grupo Local (**gpedit.msc**).
1. Navegue **até User** **ConfigurationAdministrative** >  **TemplatesMicrosoft**  >  Outlook 2016  > **Miscellaneous**.
1. Selecione a **configuração Desativar Outlook web de complementos cuja COM ou VSTO add-in equivalente está instalado**.
1. Abra o link para editar a configuração de política.
1. Na caixa de **diálogo Outlook de web para desativar**:
    1. Definir **o nome** do valor `Id` como o encontrado no manifesto do complemento da Web. **Importante**: *Não adicione* chaves ao `{}` redor da entrada.
    1. **Desmarcar** Valor como `ProgId` o do complemento COM/VSTO equivalente.
    1. Selecione **OK** para colocar a atualização em vigor.
    ![Captura de tela mostrando a caixa de diálogo "Outlook de web para desativar".](../images/outlook-deactivate-gpo-dialog.png)

## <a name="equivalent-behavior-for-users"></a>Comportamento equivalente para usuários

Quando um complemento [COM](#specify-an-equivalent-com-add-in) equivalente é especificado, o Office no Windows não exibirá a interface do usuário do seu Office Add-in se o complemento COM equivalente estiver instalado. Office oculta apenas os botões de faixa de opções do Office e não impede a instalação. Portanto, Office seu complemento ainda aparecerá nos seguintes locais dentro da interface do usuário.

- Em **Meus complementos**
- Como entrada no gerenciador de faixa de opções (Excel, Word e PowerPoint somente)

> [!NOTE]
> A especificação de um complemento COM equivalente no manifesto não tem efeito em outras plataformas, como Office na Web ou no Mac.

Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o Office Add-in.

### <a name="appsource-acquisition-of-an-office-add-in"></a>Aquisição do AppSource de um Office Add-in

Se um usuário adquirir o Office do AppSource e o complemento COM equivalente já estiver instalado, Office:

1. Instale o Office de usuário.
2. Ocultar a Office interface do usuário de complemento na faixa de opções.
3. Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.

### <a name="centralized-deployment-of-office-add-in"></a>Implantação centralizada do Office Desemporto

Se um administrador implantar o Office Add-in em seu locatário usando a implantação centralizada e o complemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver quaisquer alterações. Depois Office reiniciar, ele irá:

1. Instale o Office de usuário.
2. Ocultar a Office interface do usuário de complemento na faixa de opções.
3. Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.

### <a name="document-shared-with-embedded-office-add-in"></a>Documento compartilhado com o Office Incorporado

Se um usuário tiver o add-in COM instalado e, em seguida, receber um documento compartilhado com o Office Add-in incorporado, ao abrir o documento, Office irá:

1. Solicitar que o usuário confie no Office Add-in.
2. Se for confiável, o Office de usuário será instalado.
3. Ocultar a Office interface do usuário de complemento na faixa de opções.

## <a name="other-com-add-in-behavior"></a>Outro comportamento de complemento COM

### <a name="excel-powerpoint-word"></a>Excel, PowerPoint, Word

Se um usuário desinstalar o add-in COM equivalente, Office em Windows restaurará a interface do usuário do Office Add-in.

Depois de especificar um add-in COM equivalente para seu Office de Office, o Office interrompe o processamento de atualizações para seu Office Add-in. Para adquirir as atualizações mais recentes para o Office, o usuário deve primeiro desinstalar o complemento COM.

### <a name="outlook"></a>Outlook

O complemento COM/VSTO deve ser conectado quando Outlook é iniciado para que o complemento da Web correspondente seja desabilitado.

Se o complemento COM/VSTO for desconectado durante uma sessão de Outlook subsequente, o complemento da Web provavelmente permanecerá desabilitado até que Outlook seja reiniciado.

## <a name="see-also"></a>Confira também

- [Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário XLL](../excel/make-custom-functions-compatible-with-xll-udf.md)
