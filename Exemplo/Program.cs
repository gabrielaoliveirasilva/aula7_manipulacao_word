using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region  Criacao do documento
                // Cria um documento com o nome exemplodoc
                Document exemploDoc = new Document
                ();
            #endregion

            #region Criacao de secao no documento
                // Adiciona ima seção com o nome secaoCapa ao documento
                // Cada secao pode ser entendida como uma pagina o documento
                Section secaoCapa = exemploDoc.AddSection
                ();
            #endregion

            #region Criar um paragrafo
                // Cria um paragrafo com o nome titulo e aciona á seção secaoCapa
                // Os paragrafos são necessários para inserção de textos, imagens, tabelas etc
                Paragraph titulo = secaoCapa.AddParagraph
                ();    
            #endregion

            #region Adiciona texto ao paragrafo
                // Adiciona o texto Exemplo de titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de título\n\n");
            #endregion

            #region Formatar paragrafo
                // Através da propriedade Ho
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //
                ParagraphStyle estilo01 = new ParagraphStyle(exemploDoc);

                // Adiciona um nome ao estilo01
                estilo01.Name = "Cor do titulo";

                // Definir a cor do texto
                estilo01.CharacterFormat.TextColor = Color.DarkBlue;

                // Define que o texto será em negrito
                estilo01.CharacterFormat.Bold = true; 

                // Adiciona o estilo01 ao documento exemploDoc
                exemploDoc.Styles.Add(estilo01);

                // Aplica o estilo01 ao parágrafo titulo
                titulo.ApplyStyle(estilo01.Name);
            #endregion

            #region Trabalhar com tabulação
                // Adiciona um paragrafo textoCapa á seção secaoCapa
                Paragraph textoCapa = secaoCapa.AddParagraph();

                textoCapa.AppendText("\tEste é um exemplo de texto com tabulação\n");

                // Adiciona um novo parágrafo à mesma seção (secaoCapa)
                Paragraph textoCapa2 = secaoCapa.AddParagraph();

                // Adiociona um texto ao parágrafo textoCapa
                textoCapa2.AppendText("\tBasicamente, então, uma seção representa uma página do documento e os parágrafos dentro de uma mesma seção," + "obviamente, aparecem na mesma página");
            #endregion

            #region Inserir imagens
                // Adiciona um parágrafo à seção Capa
                Paragraph imagemCapa = secaoCapa.AddParagraph();

                // Adiciona um texto ai parágrafo imagemCapa
                imagemCapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao documento");

                // Centraliza horozontalmente a parágrafo imagemCapa
                imagemCapa.Format.HorizontalAlignment =HorizontalAlignment.Center;

                // Adiciona uma imagem com o nome imagemExemplo ao parágrafo imagemCapa
                DocPicture imagemExemplo = imagemCapa.AppendPicture(Image.FromFile(@"saida\img\logo_csharp.png"));

                // Define uma largura e uma altura para a imagem
                imagemExemplo.Width  = 300;
                imagemExemplo.Height = 300;
            #endregion

            #region Adicionar nova seção
                // Adiciona uma nova seção
                Section secaoCorpo = exemploDoc.AddSection();

                // Adiciona um parágrafo à seção secaoCorpo
                Paragraph paragraphCorpo1 = secaoCorpo.AddParagraph();

                paragraphCorpo1.AppendText("\tEste é um Exemplo de parágrafo criado em uma nova seção." + "\tComo foi criada uma nova seção, perceba que este texto aparece em uma nova página.");
                #endregion

                #region Adicionar uma tabela
                    // Adiciona uma tabela à seção secaoCorpo
                    Table tabela = secaoCorpo.AddTable(true);

                    // Cria o cabeçalho da tabela
                    String[] cabecalho = {"item","Descrição","Qtd.","preço unit.","preço"};

                    string[][] dados = {
                        new String[]{"Cenoura","Vegetal muito mutritivo","1","R$ 4,00","R$ 4,00"},

                        new String[]{"batata","Vegetal muito consumido","2","R$ 45,00","R$ 10,00"},

                        new String[]{"Alfase","Vegetal utilizado desde 500 a.C","1","R$ 1,50","R$ 1,50"},

                        new String[]{"Tomate","Tomate é uma fruta","2","R$ 6,00","R$ 12,00"},
                    };

                    // Adiciona as células na tabela
                    tabela.ResetCells(dados.Length + 1, cabecalho.Length);

                    // Adiciona uma linha da posição [0] do vetor de limites
                    // E define que esta linha é o cabecalho
                    TableRow Linha1 = tabela.Rows[0];
                    Linha1.IsHeader = true;

                    // define a altura da linha
                    Linha1.Height = 23;

                    // Formatação do cabeçalho
                    Linha1.RowFormat.BackColor = Color.AliceBlue;

                    for (int i = 0; i < cabecalho.Length; i++)
                    {
                        Paragraph p = Linha1.Cells[i].AddParagraph();
                        Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                        // Formatação dos dados do cabeçalho
                        TextRange TR = p.AppendText(cabecalho[i]);
                        TR.CharacterFormat.FontName  = "Calibr";
                        TR.CharacterFormat.FontSize  = 14;
                        TR.CharacterFormat.TextColor = Color.Teal;
                        TR.CharacterFormat.Bold      = true;
                    }

                    // Adicione as linhas do corpo da tabela
                    for (int r = 0; r < dados.Length; r++)
                    {
                        TableRow linhaDados = tabela.Rows[r + 1];

                        // Define a altura da linha
                        linhaDados.Height = 20;

                    for (int c = 0; c < dados[r].Length; c++)
                    {
                        // Alinha as células
                        linhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                        // preenche os dados nas linhas
                        Paragraph p2 = linhaDados.Cells[c].AddParagraph();
                        TextRange TR2 = p2.AppendText(dados[r][c]);

                        // Formata ad células
                        p2.Format.HorizontalAlignment = 
                        HorizontalAlignment.Center;
                        TR2.CharacterFormat.FontName = "Calibri";
                        TR2.CharacterFormat.FontSize = 12;
                        TR2.CharacterFormat.TextColor = Color.Brown;
                    }
                    }
                #endregion

                #region Salvar aequivo
                    // Salve o arquivo em .Docx
                    // utilizar o método SaveToFile para salvar o arquivo no formato desejado
                    // Assim como no Word, caso já exista um arquivo com este nome, é substituido
                    exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
                #endregion
        }
    }
}
