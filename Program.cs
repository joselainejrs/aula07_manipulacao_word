
using System;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using Spire.Doc.Fields;

namespace exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Criacao do documento
                // Cria um documento com o nome exemploDoc
                Document exemploDoc = new Document();
            #endregion

            #region Criacao de secao no documento
                //Adiciona uma seção capa com o nome secaoCapa ao documento
                Section secaoCapa = exemploDoc.AddSection();
            #endregion
                //Cria um paragrafo com o nome titulo

            #region Criar um paragrafo
                //Cria um paragrafo com o nome titulo e adiciona á secaoCapa
                //Os paragrafos são necessários para inserção de textos, imagens, tabelas etc
                Paragraph titulo = secaoCapa.AddParagraph();
            #endregion   

            #region Adiciona texto ao paragrafo
                //Adiciona o texto Exemplo de titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de titulo\n\n");
             #endregion  


             #region Formatar paragrafo
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;    
                
                //Cria um estilo com o nome estilo 01 e adiciona ao ]documento
                ParagraphStyle estilo01 = new ParagraphStyle(exemploDoc);    

                //Adiciona um nome ao estilo01
                estilo01.Name = "cor do titulo";
                
                //Definir a cor do texto
                estilo01.CharacterFormat.TextColor = Color.DarkBlue;   

                // Define que o texto será em negrito
                estilo01.CharacterFormat.Bold = true;

                // Adiciona o estilo01 ao documento exemploDOC
                exemploDoc.Styles.Add(estilo01);

                titulo.ApplyStyle(estilo01.Name);
             #endregion

             #region trabalhar com Tabulação
                //Adicione um paragrafo TextoCapa a seção secaoCapa
                Paragraph textoCapa = secaoCapa.AddParagraph();

                //Adiciona um texto ao paragrafo com tabulação
                textoCapa.AppendText("\tEste é um Exemplo de texto com Tabulação\n");

                //Adiciona um novo paragrafo á mesma seçao (secaoCapa)   
                Paragraph textoCapa2 = secaoCapa.AddParagraph();

                //Adiciona um texto ao paragrafo textoCapa2 com concatenção                
                textoCapa2.AppendText("\tBasicamente, então, uma secao representa uma pagina do documento e os paragrafos dentro de uma mesma seção," + "Obviamente, aparecem na mesma página"); 
             #endregion

             #region Insira uma Imagem
                //Adiciona um paragrafo a seção secaoCapa
                Paragraph ImagemCapa = secaoCapa.AddParagraph();

                //Adiciona um texto ao paragrafo ImagemCapa
                ImagemCapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao document\n\n");

                //Centralizado horizontamente o paragrafo e ImagemCapa
                ImagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // Adiciona um imagem com o nome imagemExemplo ao paragrafo imagemCapa
                DocPicture imagemExemplo = ImagemCapa.AppendPicture(Image.FromFile(@"saida\img\logo_csharp.png"));  

                //Define uma largura e uma altura para a imagem
                imagemExemplo.Width = 300;             
                imagemExemplo.Height = 300;          
             #endregion

             #region Adicionar a nova seção
                //adiciona uma nova seção
                Section secaoCorpo = exemploDoc.AddSection();

                //Adiciona um paragrafo a seção secaoCorpo
                Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph();

                //Adiciona um Texto ao paragrafo paragrafoCorpo1
                paragrafoCorpo1.AppendText("\t Eeste é um exemplo de paragrafo criado em uma nova seção." + "\t Como foi criada uma nova seção, perceba que este texto aparece em uma nova página.");
                #endregion

                //Adiciona uma tabela á seção secaoCorpo
                 #region Adicinar uma Tabela 
                    Table tabela = secaoCorpo.AddTable(true);

                    //Cria o cabeçalho da tabela
                    String[] cabecalho ={"Item", "Descrição", "Qtd", "Preço", " Preço Unit"};

                    String[] [] dados = {
                        new String[]{"Cenoura", "Vegetal muito Nutritvo", "1", "R$ 4,00", "R$ 4,00"},
                        new String[]{"Batata", "Vegetal muito Nutritvo", "2", "R$ 6,00", "R$ 12,00"},
                        new String[]{"Cebola", "Vegetal muito Nutritvo", "1", "R$ 4,00", "R$ 4,00"},
                        new String[]{"Beterraba", "Vegetal muito Nutritvo", "2", "R$ 4,00", "R$ 8,00"},
                    };

                        //Adicina as células na tabela
                        tabela.ResetCells(dados.Length + 1, cabecalho.Length);
                        
                        //Adicina uma linha na posição [0] do vetor de linha
                        // E define que esta linha é o cabeçalho
                        TableRow Linha1 = tabela.Rows[0];
                        Linha1.IsHeader = true;

                        //Define altura da linha
                        Linha1.Height = 23;

                        //Formatação do cabeçalho
                        Linha1.RowFormat.BackColor = Color.AliceBlue;

                        // Percorre as colunas do cabeçalho
                        for (int i = 0; i < cabecalho.Length; i++)
                        {
                            //alinhamnento das celulas
                            Paragraph p =  Linha1.Cells[i].AddParagraph();
                            Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                            p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                            //Formatação  dod dados do cabeçalhos
                            TextRange TR = p.AppendText(cabecalho[i]);
                            TR.CharacterFormat.FontName = "Calibri";
                            TR.CharacterFormat.FontSize = 14;
                            TR.CharacterFormat.TextColor = Color.Teal;
                            TR.CharacterFormat.Bold = true;
                        }

                        // Adicina as linhas do corpo da tabela
                        for (int r = 0; r < dados.Length; r++)
                        {
                            TableRow LinhaDados  = tabela.Rows[r +1];

                            //Define a açtura da linha
                            LinhaDados.Height = 20;
                        

                            for (int c = 0; c < dados[r].Length; c++)
                            {
                                //Alinha as Células
                                LinhaDados.Cells[c] .CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                                //Preenche os dados  nas linhas
                                Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                                TextRange TR2 = p2.AppendText(dados[r][c]);

                                //Formata as linhas
                                p2.Format.HorizontalAlignment =HorizontalAlignment.Center;
                                TR2.CharacterFormat.FontName = "Calibri";
                                TR2.CharacterFormat.FontSize = 12;
                                TR2.CharacterFormat.TextColor = Color.Brown; 
                        }                         
                            }
                          #endregion

                #region Salvar arquivo
                    //Salva o arquivo em Docx
                    //Utiliza o método SaveTiFile para salavar o arquvo no formato desejado
                    //Assim como no word, caso já exita um arquivo com este nome, é substituido
                    exemploDoc.SaveToFile (@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
                #endregion
        }
    }
}

