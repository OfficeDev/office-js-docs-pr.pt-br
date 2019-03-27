---
title: Suplementos de conteúdo do Office
description: Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou do PowerPoint que concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 46268f963545c3f5b7f45b9b590dc772ba37292f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870538"
---
# <a name="content-office-add-ins"></a>Suplementos de conteúdo do Office

Suplementos de conteúdo são superfícies que podem ser incorporadas diretamente em documentos do Excel ou PowerPoint. Os suplementos de conteúdo concedem aos usuários acesso a controles de interface que executam códigos para modificar documentos ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.  

*Figura 1. Layout típico dos suplementos de conteúdo*

![Imagem de exemplo exibindo um layout típico de suplementos de conteúdo.](../images/overview-with-app-content.png)

## <a name="best-practices"></a>Práticas recomendadas

- Inclua alguns elementos de navegação ou comando, como CommandBar ou Pivot, na parte superior do suplemento.
- Inclua um elemento da marca, como BrandBar, na parte inferior do suplemento (aplica-se apenas a suplementos do Excel e do PowerPoint).

## <a name="variants"></a>Variantes

Os tamanhos dos suplementos de conteúdo para Excel e PowerPoint na área de trabalho do Office e do Office 365 são especificados pelo usuário.

## <a name="personality-menu"></a>Menu de personalidade

Menus de personalidade podem obstruir elementos de navegação e comando localizados perto da parte superior direita do suplemento. Veja a seguir as dimensões atuais do menu personalidade no Windows e Mac.

No Windows, o menu de personalidade mede 12 x 32 pixels, conforme mostrado.

*Figura 2. Menu de personalidade no Windows* 

![Imagem mostrando o menu do personalidade na área de trabalho do Windows](../images/personality-menu-win.png)


No Mac, o menu de personalidade mede 26 x 26 pixels, mas flutua 8 pixels a partir da direita e 6 pixels a partir do topo, o que aumenta o espaço ocupado para 34 x 32 pixels, como mostrado.

*Figura 3. Menu de personalidade no Mac*

![Imagem mostrando o menu de personalidade na área de trabalho do Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a>Implementação

Para ver um exemplo que implementa um suplemento de conteúdo, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="support-considerations"></a>Considerações sobre o suporte

- Verifique se os suplementos do Office funcionarão em uma [plataforma de host do Office específica](/office/dev/add-ins/overview/office-add-in-availability). 
- Alguns suplementos de conteúdo podem obrigar o usuário a "confiar" nele para ler e gravar no Excel ou PowerPoint. Você pode declarar no manifesto do suplemento quais [níveis de permissão](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) deseja que o usuário tenha.  
- Os suplementos de conteúdo são compatíveis com o Excel e PowerPoint nas versões do Office 2013 e posteriores. Se você abrir um suplemento em uma versão do Office não compatível com os suplementos web do Office, eles aparecerão como imagem.

## <a name="see-also"></a>Confira também

- [Disponibilidade de host e plataforma para suplementos do Office](/office/dev/add-ins/overview/office-add-in-availability)
- [Office UI Fabric em Suplementos do Office](/office/dev/add-ins/design/office-ui-fabric)
- [Padrões de design da experiência do usuário para suplementos do Office](/office/dev/add-ins/design/ux-design-pattern-templates)
- [Solicitar permissões para uso da API em suplementos do painel de tarefas e conteúdo](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
