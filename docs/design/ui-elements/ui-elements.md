# <a name="office-add-in-ui-elements"></a>Elementos da interface do usuário de Suplemento do Office

Você pode usar dois tipos de elementos de interface do usuário nos seus Suplementos do Office: 

- Comandos de suplemento 
- Interfaces baseadas em HTML personalizadas

## <a name="add-in-commands"></a>Comandos de suplemento
Os comandos são definidos no [manifesto XML do suplemento](../../../docs/develop/define-add-in-commands.md) e são renderizados como extensões UX nativas na interface do usuário do Office. Por exemplo, você pode usar comandos de suplemento para adicionar botões à faixa de opções do Office. 

![Uma imagem mostrando comandos do suplemento e elementos personalizados da interface do usuário HTML em um suplemento](../../images/layouts_addInCommands_v0.03.png)

Atualmente, suplemento comandos só há suporte para suplementos do email. Para saber mais, consulte [comandos do suplemento para email](../../outlook/add-in-commands-for-outlook.md). 

O Excel, o PowerPoint e o Word têm pontos de entrada predefinidos para suplementos de conteúdo e painel de tarefas na guia Inserir da faixa de opções do Office. A funcionalidade de comando personalizado para suplementos do painel de tarefas e conteúdo estará disponível em breve. 

![Uma imagem que mostra a guia Inserir da faixa de opções do Word](../../images/Word-insert-tab.png)

## <a name="custom-html-based-ui"></a>Interface de usuário personalizada baseada em HTML
Os suplementos podem incorporar interfaces de usuários personalizadas baseadas em HTML em clientes do Office. Os contêineres que estão disponíveis para exibir a interface do usuário variam de acordo com o tipo de suplemento. Por exemplo, os suplementos do painel de tarefas exibem interfaces de usuários personalizadas baseadas em HTML no painel à direita do documento; os suplementos de conteúdo exibem a interface do usuário personalizada diretamente no documento do Office.

Independentemente do tipo de suplemento criado, você pode usar blocos de construção comuns para criar uma interface do usuário personalizada baseada em HTML. Recomendamos usar o [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) para esses elementos da interface do usuário para que seu suplemento se integre à aparência do Office. Sinta-se à vontade para usar também seus próprios elementos de interface do usuário para expressar sua marca.

O Office UI Fabric fornece os seguintes elementos de interface do usuário:

- Tipografia
- Cor
- Ícones
- Animações
- Componente de entrada
- Layouts
- Elementos de navegação

Você pode baixar o [Office UI Fabric do Github](https://github.com/OfficeDev/Office-UI-Fabric).

Para um exemplo que mostra como usar o Office UI Fabric em suplementos, confira [Exemplo de suplemento do Office UI Fabric](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample).

**Observação:** Se você decidir usar um conjunto de fontes e ícones personalizados, verifique se eles não entram em conflito com os do Office. Por exemplo, não use ícones que são iguais ou semelhantes aos do Office, mas que façam a diferença no seu suplemento. 

### <a name="creating-a-customized-color-palette"></a>Criando uma paleta de cores personalizadas
Se você decidir usar sua própria paleta de cores, considere o seguinte: 
 
- Use as cores para ajudar a comunicar o valor da sua marca aos usuários, e para adicionar emoção e prazer à experiência do usuário do seu suplemento.
- Use as cores de forma significativa e consistente no seu suplemento. Por exemplo, escolha uma cor como característica para dar a seu suplemento um tema visual consistente.
- Evite usar a mesma cor para elementos interativos e não interativos. Se você usar cores para indicar itens com os quais os usuários podem interagir, como navegação, links e botões, não use a mesma cor para itens estáticos.
- Se você usar uma cor para texto ou texto branco em uma tela de fundo colorido, verifique se as cores têm contraste suficiente para atender às diretrizes de acessibilidade (índice de contraste 4.5:1).
- Lembre-se do daltonismo, use mais do que apenas cores para indicar interatividade.

### <a name="theming"></a>Temas 
Não importa se você deseja adotar o esquema de cores do Office ou usar seu próprio esquema de cores, recomendamos a você usar nossas APIs de Temas. Os suplementos que fazem parte da experiência de temas do Office parecerão muito mais integrados ao Office.


- Para suplementos do painel de tarefas e email, use a propriedade [Context.officeTheme](http://dev.office.com/reference/add-ins/shared/office.context.officetheme) para combinar o tema dos aplicativos do Office. Atualmente, essa API só está disponível no Office 2016.  
- Para suplementos de conteúdo do PowerPoint, confira [Usar os temas do Office em seus suplementos do PowerPoint](../../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md).

<!-- Link to theming API docs and Humberto's seed sample. Add screenshot of themed add-in. -->



