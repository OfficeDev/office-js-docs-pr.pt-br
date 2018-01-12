# <a name="deploy-and-publish-your-office-add-in"></a>Implantar e publicar seu suplemento do Office

Você pode usar um dos vários métodos para implantar o suplemento do Office para teste ou distribuição aos usuários.

|**Method**|**Use...**|
|:---------|:------------|
|[Sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Usado como parte do processo de desenvolvimento para testar o suplemento em execução no Windows, Office Online, iPad ou Mac.|
|[Implantação Centralizada](centralized-deployment.md)|Em uma implantação híbrida ou de nuvem para distribuir seu suplemento aos usuários na sua organização usando o centro de administração do Office 365.|
|[Catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Usado para distribuir o suplemento aos usuários da organização em um ambiente local.|
|[Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store)|Usado para distribuir o suplemento publicamente aos usuários.|
|[Servidor Exchange](#outlook-add-in-deployment)|Usado para distribuir suplementos do Outlook aos usuários em um ambiente local ou online.|

>**Observação:** caso pretenda enviar o suplemento para a Office Store, verifique se você está em conformidade com as [políticas de validação da Office Store](https://msdn.microsoft.com/pt-BR/library/jj220035.aspx). Por exemplo, para passar na validação, o suplemento deve funcionar em todas as plataformas com suporte para os métodos definidos. Saiba mais na [seção 4.12](https://dev.office.com/officestore/docs/validation-policies#4-apps-and-add-ins-behave-predictably) e na [página de hospedagem e disponibilidade do Suplemento do Office](https://dev.office.com/add-in-availability).

## <a name="deployment-options-by-office-host"></a>Opções de implantação pelo host do Office

As opções de implantação disponíveis dependem do host do Office que você pretende usar e do tipo de suplemento que você pretende criar.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Opções de implantação de suplementos para Word, Excel e PowerPoint

| Ponto de extensão | Sideloading | Centro de administração do Office 365 |Office Store| Catálogo do SharePoint*  |
|:----------------|:------------|:-------------------|:--------------------------------|:-------------|
| Conteúdo         | X           | X                  | X                               | X|
| Painel de tarefas       | X           | X                  | X                               | X|
| Comando         | X           | X                  | X                               |  |

&#42; Os catálogos do SharePoint não são compatíveis com o Office 2016 para Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Opções de implantação para suplementos do Outlook

| Ponto de extensão | Sideloading | Servidor Exchange | Office Store |
|:---------|:------------|:----------------|:-------------|
| Aplicativo de email | X           | X               | X            |
| Comando  | X           | X               | X            |

## <a name="deployment-methods"></a>Métodos de implantação

As seções a seguir fornecem informações adicionais sobre os métodos de implantação mais comumente usados para distribuir suplementos do Office para usuários da organização.

Saiba mais sobre como os usuários finais podem adquirir, inserir e executar suplementos em [Começar a usar seu Suplemento do Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>Implantação Centralizada por meio do centro de administração do Office 365 

No centro de administração do Office 365, fica mais fácil para o administrador implantar Suplementos do Office para usuários e grupos dentro da organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. Você pode usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

Confira mais informações em [Publicar Suplementos do Office usando a Implantação Centralizada por meio do Centro de Administração do Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>Implantação de catálogo do SharePoint

O catálogo de suplementos do SharePoint é uma coleção de sites especial que você pode criar para hospedar suplementos dos aplicativos Word, Excel e PowerPoint. Como os catálogos do SharePoint não oferecem suporte para os novos recursos de suplemento implementados no nó `VersionOverrides` do manifesto, inclusive comandos do suplemento, recomendamos usar a Implantação Centralizada por meio do centro de administração, se possível. Por padrão, os comandos do suplemento implantados por meio do catálogo do SharePoint abrem em um painel de tarefas.

Se você está implantando suplementos em um ambiente local, use um catálogo do SharePoint. Para saber mais, confira, [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

>**Observação:** Os catálogos do SharePoint não são compatíveis com o Office 2016 para Mac. Para implantar Suplementos do Office em clientes do Mac, você deve enviá-los para a [Office Store]. 

### <a name="outlook-add-in-deployment"></a>Implantação de suplemento do Outlook

Em relação aos ambientes locais e online que não usam o serviço de identidade do Microsoft Azure AD, é possível implantar suplementos do Outlook por meio do servidor Exchange. 

Requisitos de implantação de suplemento do Outlook:

- Office 365, Exchange Online ou Exchange Server 2013 ou posterior
- Outlook 2013 ou posterior

Para atribuir suplementos a locatários, use o Centro de administração do Exchange para carregar o manifesto diretamente de um arquivo ou de uma URL ou para adicionar um suplemento da Office Store. Para atribuir suplementos a usuários individuais, é necessário usar o Exchange PowerShell. Para saber mais, confira o artigo [Instalar ou remover suplementos do Outlook para a organização](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) no TechNet.

## <a name="additional-resources"></a>Recursos adicionais

- [Implantar e instalar suplementos do Outlook para teste](../outlook/testing-and-tips.md) 
- [Enviar para a Office Store][Office Store]
- [Diretrizes de design para suplementos do Office](../design/add-in-design)
- [Criar suplementos eficientes para a Office Store](https://msdn.microsoft.com/pt-BR/library/jj635874.aspx)
- [Solucionar erros de usuários com suplementos do Office](../testing/testing-and-troubleshooting.md)

[Office Store]: http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx
[Office Add-in host and platform availability]: http://dev.office.com/add-in-availability
