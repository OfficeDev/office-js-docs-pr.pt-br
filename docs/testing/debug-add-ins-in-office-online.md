---
title: Depurar suplementos no Office Online
description: Como usar o Office Online para testar e depurar seus suplementos.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ff77f3d8b3e332288d4ccb3e2d2305d1b1c4a825
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871147"
---
# <a name="debug-add-ins-in-office-online"></a>Depurar suplementos no Office Online


Você pode criar e depurar suplementos em um computador que não esteja executando o Windows ou os clientes de área de trabalho do Office 2013 ou do Office 2016, por exemplo, se você estiver desenvolvendo no Mac. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos. 

## <a name="prerequisites"></a>Pré-requisitos

Para começar:

- Obtenha uma conta de desenvolvedor do Office 365, se já não tiver uma, ou o acesso a um site do SharePoint.
    
  > [!NOTE]
  > Para se inscrever para uma assinatura gratuita de desenvolvedor do Office 365, ingresse no [Programa de Desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program). Confira as instruções passo a passo sobre como participar do Programa para Desenvolvedores do Office 365, entrar e configurar sua assinatura na [documentação do Programa para Desenvolvedores do Office 365](/office/developer-program/office-365-developer-program).
     
- Configure um catálogo de suplementos no Office 365 (SharePoint Online). Um catálogo de suplementos é um conjunto de sites dedicado no SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office. Se você tiver seu próprio site do SharePoint, pode configurar uma biblioteca de documentos do catálogo de suplementos. Para saber mais, confira [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo de suplementos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Depurar seu suplemento do Excel Online ou do Word Online

Para depurar seu suplemento usando o Office Online:

1. Implante o suplemento em um servidor que dê suporte a SSL.
    
    > [!NOTE]
    > Recomendamos que você use o [gerador Yeoman](https://github.com/OfficeDev/generator-office) para criar e hospedar seu suplemento.
     
2. No seu [arquivo de manifesto de suplemento](../develop/add-in-manifests.md), atualize o valor do elemento **SourceLocation** para incluir um URI absoluto, em vez de relativo. Por exemplo:
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Carregue o manifesto na biblioteca de Suplementos do Office no catálogo de suplementos no SharePoint.
    
4. Inicie o Excel Online ou o Word Online do inicializador de aplicativos no Office 365 e abra um novo documento.
    
5. Na guia Inserir, escolha **Meus Suplementos** ou **Suplementos do Office** para inserir seu suplemento e testá-lo no aplicativo.
    
6. Use seu depurador de navegador favorito para depurar o suplemento.

## <a name="potential-issues"></a>Possíveis problemas    

A seguir apresentamos alguns problemas que você pode encontrar ao depurar:
    
- Alguns erros de JavaScript que você vê podem vir do Office Online.
      
- O navegador pode mostrar um erro de certificado inválido que você deve ignorar.
      
- Se você definir pontos de interrupção no seu código, o Office Online pode lançar uma mensagem de erro indicando que não é possível salvar.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Políticas de validação do AppSource](/office/dev/store/validation-policies)  
- [Criar aplicativos e suplementos eficazes para o AppSource](/office/dev/store/create-effective-office-store-listings)  
- [Solucionar erros de usuários com suplementos do Office](testing-and-troubleshooting.md)
    
