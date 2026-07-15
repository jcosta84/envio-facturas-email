package cv.edu.us.envio;

import java.util.Iterator;
import java.util.NoSuchElementException;

/**
 * Estrutura de dados própria utilizada para guardar resultados de envio.
 */
public class ListaLigada<T> implements Iterable<T> {
    private No<T> inicio;
    private No<T> fim;
    private int tamanho;

    private static final class No<T> {
        private final T valor;
        private No<T> proximo;

        private No(T valor) {
            this.valor = valor;
        }
    }

    public void adicionar(T valor) {
        No<T> novo = new No<>(valor);
        if (inicio == null) {
            inicio = fim = novo;
        } else {
            fim.proximo = novo;
            fim = novo;
        }
        tamanho++;
    }

    public int tamanho() {
        return tamanho;
    }

    public boolean estaVazia() {
        return tamanho == 0;
    }

    @Override
    public Iterator<T> iterator() {
        return new Iterator<>() {
            private No<T> atual = inicio;

            @Override
            public boolean hasNext() {
                return atual != null;
            }

            @Override
            public T next() {
                if (atual == null) throw new NoSuchElementException();
                T valor = atual.valor;
                atual = atual.proximo;
                return valor;
            }
        };
    }
}
