---
title: Atualize para a biblioteca mais recente da API JavaScript para Office e esquema de manifesto de suplemento versão 1.1
description: Atualize os arquivos JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto do suplemento no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 12/04/2017
ms.openlocfilehash: 676d1cde832399b2518a6393c38e7c4bf78d608c
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505759"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>Atualize para a biblioteca mais recente da API JavaScript para Office e esquema de manifesto de suplemento versão 1.1

Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.

## <a name="use-the-most-up-to-date-project-files"></a>Usar os arquivos de projeto mais atualizados

Se você usar o Visual Studio para desenvolver seu suplemento, para usar os [membros de API mais recentes](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office?view=office-js) da API JavaScript para Office e os [recursos da v1.1 do manifesto de suplemento](../develop/add-in-manifests.md) (que é validado em relação a offappmanifest-1.1.xsd), é preciso baixar e instalar o [Visual Studio 2015 e a versão mais recente do Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs).

Se você usar um editor de texto ou IDE diferente do Visual Studio para desenvolver seu suplemento, será necessário atualizar as referências ao CDN para Office.js e a versão do esquema referenciada no manifesto do seu suplemento.

Para executar um suplemento desenvolvido usando as APIs novas e atualizadas do Office.js e os recursos de manifesto de suplementos, seus clientes devem estar executando produtos locais do Office 2013 SP1 ou versão posterior e, quando aplicável, o SharePoint Server 2013 SP1 e produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou os produtos hospedados online equivalentes: Office 365, SharePoint Online e Exchange Online.

Para baixar os produtos Office, SharePoint e Exchange SP1, consulte o seguinte:

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos de desktop relacionados](http://support.microsoft.com/kb/2850036)
    
- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos de servidor relacionados](http://support.microsoft.com/kb/2850035)
    
- [Descrição do Exchange Server 2013 Service Pack 1](http://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Atualizando um projeto de suplemento do Office criado com o Visual Studio

Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e o esquema de manifesto do suplemento, é possível atualizar os arquivos de um projeto usando o **Gerenciador de Pacotes NuGet** e, em seguida, atualizar as páginas em HTML do suplemento para fazer referência a eles. 

Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para a versão mais recente


1. No Visual Studio 2015, abra ou crie um novo projeto de **Suplemento do Office**.
    
      - No painel à esquerda, escolha **Atualizar** e conclua o processo de atualização do pacote.
    
      - Vá para a Etapa 6.
    
2. Escolha **Ferramentas** > **Gerenciador de Pacotes NuGet** > **Gerenciar Pacotes Nuget para a solução**.
    
3. No **Gerenciador de Pacotes NuGet**, escolha **nuget.org** como a **Origem do pacote** e **Atualização disponível** como **Filtro** e selecione Microsoft.Office.js.
    
4. No painel à esquerda, escolha **Atualizar** e conclua o processo de atualização do pacote.
    
5. Na marca **head** das páginas HTML do seu suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > O `/1/` na frente de `office.js` na URL da CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.   


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualize o arquivo de manifesto no seu projeto para usar a versão 1.1 do esquema

No seu arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor da versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos elementos [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts?view=office-js) e [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host?view=office-js) ou pelos  [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE

Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e do esquema de manifesto de suplemento, é preciso atualizar as páginas HTML do suplemento para fazerem referência à CDN da biblioteca v1.1 e atualizar o arquivo de manifesto do suplemento para usar a v1.1 do esquema. 

O processo de atualização é aplicado _por projeto_ - você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

Você não precisa de cópias locais dos arquivos da API JavaScript para Office (Office.js e arquivos .js específicos do aplicativo) para desenvolver um suplemento do Office (a referência à CDN para Office.js baixa os arquivos necessários em tempo de execução). Porém, se desejar uma cópia local dos arquivos da biblioteca, pode usar o [Utilitário de linha de comando NuGet](http://docs.nuget.org/consume/installing-nuget) e o comando `Install-Package Microsoft.Office.js` para baixá-los.

> [!NOTE] 
> Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema de manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para usar a versão mais recente

1. Abra as páginas HTML do suplemento no editor de texto ou IDE.
    
2. Na marca **head** das páginas HTML do seu suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > O `/1/` na frente de `office.js` na URL da CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualize o arquivo de manifesto no seu projeto para usar a versão 1.1 do esquema

No seu arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor da versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos elementos [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts?view=office-js) e [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host?view=office-js) ou pelos [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).
    

## <a name="see-also"></a>Confira também

- [Especificar requisitos de API e hosts do Office](specify-office-hosts-and-api-requirements.md) 
- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)    
- [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)   
- [Referência de esquema para manifestos de suplementos do Office (v1.1)](../develop/add-in-manifests.md)
    
