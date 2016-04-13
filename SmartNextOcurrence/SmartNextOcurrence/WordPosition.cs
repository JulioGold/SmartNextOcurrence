using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartNextOcurrence
{
    public class WordPosition
    {
        public string Word { get; set; }

        public int Start { get; set; }

        public int End { get; set; }

        public static WordPosition PreviousWord(string text, int init)
        {
            // TODO: Take a way to this mess, take like a NextWord method

            string palavra = String.Empty;
            bool achouLetra = false;

            for (int i = init; i > 0; i--)
            {
                string letra = text.Substring(i, 1);

                #region Desativado

                //// Se é algum destes caracteres
                //if (".,:;/?(){}[]-=+*|\\_'".Contains(letra))
                //{
                //    //WordPosition wordPosition = new WordPosition() { Word = text.Substring(i + 1, init - i), Start = i + 1, End = init - i };

                //    //// Se tem um espaço logo depois, ou seja a direita, então pego o próprio caracter
                //    //if (" ".Contains(text.Substring(i + 1, init - i)))
                //    //{
                //    //    wordPosition.Word = text.Substring(i, init - i);
                //    //    wordPosition.Start = i;
                //    //    wordPosition.End = init - i;
                //    //}

                //    //return wordPosition;

                //    string word = text.Substring(i, 1);

                //    return new WordPosition()
                //    {
                //        Word = word,
                //        Start = i,
                //        End = i + word.Length
                //    };
                //}

                #endregion

                if (achouLetra && " ".Contains(letra))
                {
                    string word = text.Substring(i, (init - i));

                    return new WordPosition()
                    {
                        Word = word,
                        Start = i,
                        End = i + word.Length
                    };
                }

                if ("abcdefghijklmnopqrstuwvxyz0123456789ABCDEFGHIJKLMNOPQRSTUWVXYZ".Contains(letra))
                {
                    achouLetra = true;
                }
            }

            // Se chegou aqui então a posição é zero
            return new WordPosition() { Word = String.Empty, Start = 0, End = init };
        }

        public static WordPosition NextWord(string text, int init)
        {
            char[] delimiterChars = { ' ', ',', '.', ':', '\t', '(', ')', '{', '}', '[', ']' };

            WordPosition[] wordsPosition = WordPosition.Parse(text, delimiterChars);

            int index = 0;

            for (int i = 0; i < wordsPosition.Length; i++)
            {
                WordPosition item = wordsPosition[i];

                if (init >= item.Start && init <= item.End)
                {
                    index = i;
                    break;
                }
            }

            // Próxima palavra | Se o índice já for o último, a próxima palavra é a última
            WordPosition wordPosition = wordsPosition[(wordsPosition.Length == (index + 1)) ? index : (index + 1)];

            return wordPosition;
        }

        private static WordPosition[] Parse(string text, char[] separator)
        {
            List<WordPosition> lista = new List<WordPosition>();
            int inicio = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char letra = text.Substring(i, 1).ToCharArray()[0];

                if (separator.Contains(letra))
                {
                    string word = text.Substring(inicio, i - inicio);

                    lista.Add(
                        new WordPosition
                        {
                            Word = word,
                            Start = inicio,
                            End = i
                        });

                    inicio = i;
                }
            }

            return lista.ToArray();
        }

        public static WordPosition GetWordByPosition(WordPosition[] words, int position)
        {
            WordPosition wordPosition = new WordPosition();

            foreach (var item in words)
            {
                if (position > item.Start && position < item.End)
                {
                    wordPosition = item;
                    break;
                }
            }

            return wordPosition;
        }
    }
}
