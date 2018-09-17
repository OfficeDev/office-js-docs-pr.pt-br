---
title: Instalar a última versão do Office
description: Informações sobre como aceitar para obter as versões mais recentes do Office.
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943877"
---
# <a name="install-the-latest-version-of-office"></a>Instalar a última versão do Office

Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas versões do Office. 

## <a name="opt-in-to-getting-the-latest-builds"></a>Aceitar para receber as versões mais recentes

Aceitar para receber as versões mais recentes do Office: 

- Se você é assinante do Office 365 Home, Personal ou University, confira [Ser um Office Insider](https://products.office.com/office-insider).
- Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Se você estiver executando o Office em um Mac:
    - Inicie um programa do Office para Mac.
    - Selecione **Verificar Atualizações** no menu Ajuda.
    - Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider. 

## <a name="get-the-latest-build"></a>Obter a versão mais recente

Para obter as versões mais recentes do Office: 

1. Baixe a [Ferramenta de Implantação do Office](https://www.microsoft.com/download/details.aspx?id=49117). 
2. Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.
3. Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Execute o seguinte comando como administrador:  `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > O comando pode demorar muito para ser executado sem indicar o progresso.

Quando o processo de instalação for concluído, você terá os aplicativos do Office mais recentes instalados. Para verificar se você tem a versão mais recente, vá para o **Arquivo** > **Conta** de qualquer aplicativo do Office. Em atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office

Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:

- [Conjuntos de requisitos da API JavaScript do Word](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [Conjuntos de requisitos da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [Conjuntos de requisitos da API JavaScript do OneNote](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [Conjuntos de requisitos da API de caixa de diálogo](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Conjuntos de requisitos de API comum do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
