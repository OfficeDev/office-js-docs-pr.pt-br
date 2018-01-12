# <a name="office-add-in-design-guidelines"></a>Diretrizes de design de suplementos do Office

Aprimore a experiência do usuário no suplemento do Office desenvolvendo uma interface do usuário que corresponda à voz do Office e aplique as diretrizes de acessibilidade para garantir que o suplemento seja acessível a todos os usuários.

Se você planeja tornar seu suplemento [disponível na Office Store](https://dev.office.com/officestore/docs/submit-to-the-office-store), certifique-se de que a linguagem e o conteúdo estejam compatíveis com as [Políticas de validação](https://dev.office.com/officestore/docs/validation-policies).

## <a name="voice-guidelines"></a>Diretrizes de voz 

Ao desenvolver seus Suplementos do Office, considere o tom que você utiliza nos elementos e no texto da interface do usuário. Procure manter o tom da interface de usuário do Office, que é coloquial, envolvente e acessível aos usuários. 

Para alinhar seu texto aos princípios do tom do Office:

- **Use um estilo natural.** Escreva da maneira como você fala. Evite jargões e frases ou palavras muito técnicas. Use termos que sejam familiares aos usuários.
- **Use uma linguagem simples e direta.** Use palavras e frases curtas, e a voz ativa no seu texto. 
- **Seja consistente.** Use sempre as mesmas palavras para os mesmos conceitos.
- **Envolva os usuários.** Use o pronome “você” para falar com os usuários. Evite usar a terceira pessoa. Use imperativos para as tarefas do usuário.
- **Seja prestativo e empático.** Torne seu texto positivo, gentil, solidário e estimulante. Enfatize o que os usuários podem conseguir, não o que eles não podem.
- **Conheça seus clientes.** Leve em consideração as questões culturais e a globalização ao usar expressões idiomáticas e coloquialismos.

## <a name="accessibility-guidelines"></a>Diretrizes de acessibilidade

À medida que você projeta e desenvolve seus suplementos do Office, convém verificar se todos os usuários e clientes potenciais são capazes de usar seu suplemento com êxito. Aplique as seguintes diretrizes para garantir que sua solução seja acessível a todos os públicos.

### <a name="design-for-multiple-input-methods"></a>Projetar para vários métodos de entrada

- Certifique-se de que os usuários possam realizar operações usando apenas o teclado. Os usuários devem conseguir se mover para todos os elementos acionáveis da página usando uma combinação das teclas Tab e de setas.
- Em um dispositivo móvel, quando os usuários operam um controle por toque, o dispositivo deve fornecer um feedback sonoro útil.
- Forneça rótulos úteis para todos os controles interativos. 

### <a name="make-your-add-in-easy-to-use"></a>Tornar seu suplemento fácil de usar

- Não dependa de um único atributo, como cor, tamanho, forma, local, orientação ou som, para atribuir significados na sua interface do usuário.
- Evite alterações inesperadas de contexto, como mover o foco para outro elemento da interface do usuário sem uma ação do usuário.
- Ofereça uma maneira de verificar, confirmar ou reverter todas as ações de associação.
- Forneça uma maneira de pausar ou parar mídias, como áudio e vídeo.
- Não estabeleça um limite de tempo para uma ação do usuário.

### <a name="make-your-add-in-easy-to-see"></a>Deixar seu suplemento fácil de ver

- Evite mudanças de cor inesperadas.
- Forneça informações significativas e em tempo hábil para descrever elementos de interface do usuário, títulos e cabeçalhos, entradas e erros. Verifique se os nomes dos controles descrevem adequadamente o objetivo do controle.
- Siga as [diretrizes padrão](http://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) de contraste de cor.

### <a name="account-for-assistive-technologies"></a>Incluir tecnologias adaptativas

- Evite usar recursos que interfiram em tecnologias adaptativas, incluindo em interações visuais, auditivas ou outras.
- Não forneça o texto em um formato de imagem. Os leitores de tela não podem ler o texto em imagens.
- Forneça uma maneira para os usuários ajustarem ou desativarem todas as fontes de áudio.
- Forneça uma maneira para os usuários ativarem legendas ou descrições de áudio com fontes de áudio.
- Forneça alternativas para o som como um meio para alertar os usuários, como indicações visuais ou vibrações.

### <a name="accessibility-resources"></a>Recursos de acessibilidade

- [Diretrizes de Acessibilidade para Conteúdo da Web (WCAG) 2.0](http://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [Orientações sobre a Aplicação das WCAG 2.0 para Tecnologias de Comunicação e Informação que não Sejam da Web (WCAG2ICT)](http://www.w3.org/TR/wcag2ict/)
- [Padrão Europeu para requisitos de acessibilidade para Tecnologias de Comunicação e Informação (ICT)](http://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 



