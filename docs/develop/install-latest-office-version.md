---
title: Instalar a última versão do Office 2016
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 98dc69a7971a94b96bc3f7304fc7905f31013a87
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925231"
---
# <a name="install-the-latest-version-of-office-2016"></a>Instalar a última versão do Office 2016

Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Aceitar para receber as versões mais recentes

Aceitar para receber as versões mais recentes do Office 2016: 

- Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).
- Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Se você estiver executando o Office 2016 em um Mac:
    - Inicie um programa do Office 2016 para Mac.
    - Selecione **Verificar Atualizações** no menu Ajuda.
    - Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider. 

## <a name="get-the-latest-build"></a>Para a versão mais recente:

Para receber as versões mais recentes do Office 2016: 

1. Baixe a [Ferramenta de Implantação do Office 2016](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.
3. Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Execute o seguinte comando como administrador: `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > O comando pode demorar muito para ser executado sem indicar o progresso.

Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office 2016 instalados. Para verificar se você tem a última compilação, vá para **Arquivo**  >  **Conta** em qualquer aplicativo do Office. Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office

Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:

- [Conjuntos de requisitos da API JavaScript do Word](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets)
- [Conjuntos de requisitos da API JavaScript do Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets)
- [Conjuntos de requisitos da API JavaScript do OneNote](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets)
- [Conjuntos de requisitos da API de caixa de diálogo](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets)
- [Conjuntos de requisitos de API comum do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
