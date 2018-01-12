# <a name="install-the-latest-version-of-office-2016"></a>Instalar a última versão do Office 2016

Novos recursos de desenvolvedor, inclusive os que ainda estão na visualização, são fornecidos primeiro aos assinantes que aceitam obter as últimas compilações do Office. Para aceitar obter as últimas compilações do Office 2016: 

- Se você for assinante do Office 365 Home, Personal ou University, consulte [Ser um Office Insider](https://products.office.com/en-us/office-insider).
- Se você for um cliente corporativo do Office 365, confira [Instalar a versão de Primeiro Lançamento do Office 365 para clientes corporativos](https://support.office.com/en-us/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US).
- Se você estiver executando o Office 2016 em um Mac:
    - Inicie um programa do Office 2016 para Mac.
    - Selecione **Verificar Atualizações** no menu Ajuda.
    - Na caixa Microsoft AutoUpdate, marque a caixa para participar do programa Office Insider. 

Para obter a versão mais recente: 

1. Baixe a [Ferramenta de Implantação do Office 2016](https://www.microsoft.com/en-us/download/details.aspx?id=49117). 
2. Execute a ferramenta. Isso extrai estes dois arquivos: Setup.exe e configuration.xml.
3. Substitua o arquivo configuration.xml pelo [Arquivo de Configuração do Primeiro Lançamento](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Execute o seguinte comando como administrador:  `setup.exe /configure configuration.xml` 

>**Observação:** o comando pode demorar muito para ser executado sem indicar o progresso.

Quando o processo de instalação for concluído, você terá os últimos aplicativos do Office 2016 instalados. Para verificar se você tem a última compilação, vá para **Arquivo**  >  **Conta** em qualquer aplicativo do Office. Em Atualizações do Office, você verá o rótulo (Office Insiders) acima do número de versão.

![Uma captura de tela que mostra informações do produto com o rótulo Office Insiders](../../images/officeinsider.PNG)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Builds mínimos do Office para conjuntos de requisitos de API JavaScript para Office

Para saber mais sobre os builds mínimos de produtos para cada plataforma dos conjuntos de requisitos de API, confira o seguinte:

- [Conjuntos de requisitos de API JavaScript do Word](../../reference/requirement-sets/word-api-requirement-sets.md)
- [Conjuntos de requisitos de API JavaScript do Excel](../../reference/requirement-sets/excel-api-requirement-sets.md)
- [Conjuntos de requisitos de API JavaScript do OneNote](../../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Conjuntos de requisitos da Dialog API](../../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Conjuntos de requisitos da API do Office](../../reference/requirement-sets/office-add-in-requirement-sets.md)
