
# <a name="debug-add-ins-in-office-online"></a>Depurar suplementos no Office Online


Você pode criar e depurar suplementos em um computador que não esteja executando o Windows ou os clientes de área de trabalho do Office 2013 ou do Office 2016, por exemplo, se você estiver desenvolvendo no Mac. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos. 

Introdução:


- Obtenha uma conta de desenvolvedor do Office 365, se já não tiver uma, ou tenha acesso a um site do SharePoint.
    
     >**Observação**  Para se inscrever para uma conta gratuita de desenvolvedor do Office 365, ingresse em nosso [Programa de desenvolvedor do Office 365 ](https://dev.office.com/devprogram).
     
- Configure um catálogo de suplementos no Office 365 (SharePoint Online). Um catálogo de suplementos é um conjunto de sites dedicado no SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office. Se você tiver seu próprio site do SharePoint, pode configurar uma biblioteca de documentos do catálogo de suplementos. Para saber mais, confira [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo de suplementos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Depurar seu suplemento do Excel Online ou do Word Online

Para depurar seu suplemento usando o Office Online:


1. Implante o suplemento em um servidor que dê suporte a SSL.
    
     >**Observação:**  recomendamos que você use o [gerador Yeoman](https://github.com/OfficeDev/generator-office) para criar e hospedar seu suplemento.
     
2. No seu [arquivo de manifesto de suplemento](../../docs/overview/add-in-manifests.md), atualize o valor do elemento **SourceLocation** para incluir um URI absoluto, em vez de relativo. Por exemplo:
    
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. Carregue o manifesto na biblioteca de Suplementos do Office no catálogo de suplementos no SharePoint.
    
4. Inicie o Excel Online ou o Word Online do inicializador de aplicativos no Office 365 e abra um novo documento.
    
5. Na guia Inserir, escolha **Meus Suplementos** ou **Suplementos do Office** para inserir seu suplemento e testá-lo no aplicativo.
    
6. Use seu depurador de navegador favorito para depurar o suplemento.
    
    A seguir apresentamos alguns problemas que você pode encontrar ao depurar:
    
  - Alguns erros de JavaScript que você vê podem vir do Office Online.
    
  - O navegador pode mostrar um erro de certificado inválido que você deve ignorar.
    
  - Se você definir pontos de interrupção no seu código, o Office Online pode lançar uma mensagem de erro indicando que não é possível salvar.
    

## <a name="additional-resources"></a>Recursos adicionais


- [Práticas recomendadas para o desenvolvimento de Suplementos do Office](../overview/add-in-development-best-practices.md)
    
- [Políticas de validação para aplicativos e suplementos enviados para a Office Store (versão 1.9)](http://msdn.microsoft.com/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)
    
- [Criar aplicativos e suplementos do Office Store eficazes](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx)
    
- [Solucionar erros de usuários com os Suplementos do Office](../testing/testing-and-troubleshooting.md)
    
