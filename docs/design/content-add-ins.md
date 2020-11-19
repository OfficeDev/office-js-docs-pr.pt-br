---
title: Suplementos de conteúdo do Office
description: Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou do PowerPoint que concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: cf4ea46b4b924683756063bb36c3f2ea2b8c6764
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132078"
---
# <a name="content-office-add-ins"></a>Suplementos de conteúdo do Office

Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou PowerPoint. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.  

*Figura 1. Layout típico dos suplementos de conteúdo*

![Layout típico para suplementos de conteúdo em um aplicativo do Office](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Práticas recomendadas

- Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.
- Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento (aplica-se apenas a suplementos do Excel e do PowerPoint).

## <a name="variants"></a>Variantes

Tamanhos de suplementos de conteúdo para Excel e PowerPoint na área de trabalho do Office e o Microsoft 365 são especificados pelo usuário.

## <a name="personality-menu"></a>Menu de personalidade

Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 2. Menu de personalidade no Windows*

![menu de personalidade de 12x32-pixel na área de trabalho do Windows](../images/personality-menu-win.png)

No Mac, o menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço ocupado para 34 x 32 pixels, como mostrado.

*Figura 3. Menu de personalidade no Mac*

![menu de personalidade de 34 x 32-pixel no Mac desktop](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver um exemplo que implementa um suplemento de conteúdo, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="support-considerations"></a>Considerações sobre o suporte

- Verifique se o seu suplemento do Office funcionará em um [aplicativo ou plataforma específico do Office](../overview/office-add-in-availability.md).
- Alguns suplementos de conteúdo podem obrigar o usuário a "confiar" nele para ler e gravar no Excel ou PowerPoint. Você pode declarar no manifesto do suplemento quais [níveis de permissão](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) deseja que o usuário tenha.  
- Os suplementos de conteúdo são compatíveis com o Excel e PowerPoint nas versões do Office 2013 e posteriores. Se você abrir um suplemento em uma versão do Office não compatível com os suplementos web do Office, eles aparecerão como imagem.

## <a name="see-also"></a>Confira também

- [Disponibilidade de aplicativos e plataformas de cliente Office para Suplementos do Office](../overview/office-add-in-availability.md)
- [Office UI Fabric em Suplementos do Office](../design/office-ui-fabric.md)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
- [Solicitar permissões para uso da API em suplementos ](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
