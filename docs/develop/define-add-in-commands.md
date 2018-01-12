# <a name="define-add-in-commands-in-your-manifest"></a>Definir comandos de suplemento em seu manifesto

Os comandos de suplemento fornecem uma maneira fácil de personalizar a interface de usuário padrão do Office com os elementos de interface de usuário que executam ações; por exemplo, você pode adicionar botões personalizados na faixa de opções. Para criar comandos, adicione um nó **[VersionOverrides](../../reference/manifest/versionoverrides.md)** a um manifesto existente. 

Quando um manifesto contiver o elemento **VersionOverrides**, as versões do Word, Excel, Outlook e PowerPoint que oferecem suporte aos comandos de suplemento usarão as informações dentro desse elemento para carregá-lo. Versões anteriores de produtos do Office que não dão suporte a comandos de suplemento ignorarão o elemento.

Quando o aplicativo cliente reconhece o nó **VersionOverrides**, o nome do suplemento aparece na faixa de opções, não em um painel de tarefas no painel de leitura/composição. O suplemento não aparecerá nos dois locais.
 
## <a name="versionoverrides"></a>VersionOverrides

O elemento [VersionOverrides](../../reference/manifest/versionoverrides.md) é o elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. Há suporte no esquema manifesto v1.1 e posterior.

Há duas versões do esquema **VersionOverrides**.

| Versão do esquema | Descrição |
|----------------|-------------|
| 1.0 | Oferece suporte para comandos de suplementos para versões de área de trabalho dos aplicativos do Office. | 
| 1.1 | Adiciona suporte para [painéis de tarefas que podem ser ficados](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane) e suplementos móveis. **Observação:** atualmente, compatível apenas com o Outlook 2016 para Windows e o Outlook para iOS |

Um suplemento pode oferecer suporte a várias versões do esquema **VersionOverrides** aninhando as versões mais recentes em uma versão anterior. Isso permite que os clientes ofereçam suporte a versões mais recentes para aproveitar os novos recursos, ao mesmo tempo em que permite que os clientes mais antigos carreguem a versão mais antiga. Para saber mais, confira [Implementar várias versões](../../reference/manifest/versionoverrides.md#implementing-multiple-versions).

O elemento **VersionOverrides** inclui os seguintes elementos filho:

- [Descrição](../../reference/manifest/description.md)
- [Requirements](../../reference/manifest/requirements.md)
- [Hosts](../../reference/manifest/hosts.md)
- [Recursos](../../reference/manifest/resources.md)
- [VersionOverrides](../../reference/manifest/versionoverrides.md)

O diagrama a seguir mostra a hierarquia de elementos usada para definir comandos do suplemento. 

![Hierarquia dos elementos dos comandos de suplemento no manifesto](../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## <a name="sample-manifests"></a>Exemplos de manifestos

Para obter um exemplo de manifesto que implementa comandos de suplemento para Word, Excel e PowerPoint, confira [Exemplos de comandos de suplementos simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple).

Para obter um exemplo de manifesto que implementa comandos de suplemento para o Outlook, confira [manifesto para um suplemento do Outlook](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="additional-resources"></a>Recursos adicionais

- [Comandos de suplemento para o Outlook](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)
    
- [Manifestos de suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/manifests)
    
- [Exemplo de demonstração de comando de suplemento do Outlook](https://github.com/OfficeDev/outlook-add-in-command-demo)
