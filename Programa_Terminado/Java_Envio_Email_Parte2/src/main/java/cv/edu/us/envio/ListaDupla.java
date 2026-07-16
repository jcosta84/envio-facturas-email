package cv.edu.us.envio;

import java.util.AbstractList;
import java.util.Iterator;
import java.util.ListIterator;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.function.Predicate;

/**
 * Implementação própria de uma Lista Duplamente Ligada genérica.
 *
 * A classe estende AbstractList apenas para manter compatibilidade com
 * componentes JavaFX e APIs que recebem java.util.List, mas toda a
 * estrutura interna é implementada manualmente através de nós.
 */
public class ListaDupla<T> extends AbstractList<T> {

    private No<T> inicio;
    private No<T> fim;
    private int tamanho;

    private static final class No<T> {
        private T valor;
        private No<T> anterior;
        private No<T> proximo;

        private No(T valor) {
            this.valor = valor;
        }
    }


    public ListaDupla() {
    }

    public ListaDupla(Iterable<? extends T> valores) {
        if (valores != null) {
            for (T valor : valores) {
                add(valor);
            }
        }
    }

    @Override
    public int size() {
        return tamanho;
    }

    public int tamanho() {
        return tamanho;
    }

    public boolean estaVazia() {
        return tamanho == 0;
    }

    public void inserirInicio(T valor) {
        No<T> novo = new No<>(valor);

        if (inicio == null) {
            inicio = fim = novo;
        } else {
            novo.proximo = inicio;
            inicio.anterior = novo;
            inicio = novo;
        }

        tamanho++;
        modCount++;
    }

    public void inserirFim(T valor) {
        add(valor);
    }

    @Override
    public boolean add(T valor) {
        No<T> novo = new No<>(valor);

        if (fim == null) {
            inicio = fim = novo;
        } else {
            fim.proximo = novo;
            novo.anterior = fim;
            fim = novo;
        }

        tamanho++;
        modCount++;
        return true;
    }

    @Override
    public void add(int indice, T valor) {
        validarIndiceInsercao(indice);

        if (indice == tamanho) {
            add(valor);
            return;
        }

        if (indice == 0) {
            inserirInicio(valor);
            return;
        }

        No<T> atual = obterNo(indice);
        No<T> anterior = atual.anterior;
        No<T> novo = new No<>(valor);

        novo.anterior = anterior;
        novo.proximo = atual;
        anterior.proximo = novo;
        atual.anterior = novo;

        tamanho++;
        modCount++;
    }

    @Override
    public T get(int indice) {
        validarIndice(indice);
        return obterNo(indice).valor;
    }

    @Override
    public T set(int indice, T valor) {
        validarIndice(indice);
        No<T> no = obterNo(indice);
        T antigo = no.valor;
        no.valor = valor;
        return antigo;
    }

    @Override
    public T remove(int indice) {
        validarIndice(indice);
        No<T> no = obterNo(indice);
        return desligar(no);
    }

    @Override
    public boolean remove(Object valor) {
        No<T> atual = inicio;

        while (atual != null) {
            if (Objects.equals(atual.valor, valor)) {
                desligar(atual);
                return true;
            }
            atual = atual.proximo;
        }

        return false;
    }

    public T removerInicio() {
        if (inicio == null) {
            throw new NoSuchElementException("A lista está vazia.");
        }
        return desligar(inicio);
    }

    public T removerFim() {
        if (fim == null) {
            throw new NoSuchElementException("A lista está vazia.");
        }
        return desligar(fim);
    }

    public T pesquisar(Predicate<T> criterio) {
        Objects.requireNonNull(criterio, "O critério não pode ser nulo.");

        for (T valor : this) {
            if (criterio.test(valor)) {
                return valor;
            }
        }

        return null;
    }

    public ListaDupla<T> filtrar(Predicate<T> criterio) {
        Objects.requireNonNull(criterio, "O critério não pode ser nulo.");

        ListaDupla<T> resultado = new ListaDupla<>();

        for (T valor : this) {
            if (criterio.test(valor)) {
                resultado.add(valor);
            }
        }

        return resultado;
    }

    @Override
    public void clear() {
        No<T> atual = inicio;

        while (atual != null) {
            No<T> proximo = atual.proximo;
            atual.anterior = null;
            atual.proximo = null;
            atual.valor = null;
            atual = proximo;
        }

        inicio = null;
        fim = null;
        tamanho = 0;
        modCount++;
    }

    @Override
    public Iterator<T> iterator() {
        return new Iterator<>() {
            private No<T> atual = inicio;
            private No<T> ultimoRetornado;

            @Override
            public boolean hasNext() {
                return atual != null;
            }

            @Override
            public T next() {
                if (atual == null) {
                    throw new NoSuchElementException();
                }

                ultimoRetornado = atual;
                atual = atual.proximo;
                return ultimoRetornado.valor;
            }

            @Override
            public void remove() {
                if (ultimoRetornado == null) {
                    throw new IllegalStateException();
                }

                ListaDupla.this.desligar(ultimoRetornado);
                ultimoRetornado = null;
            }
        };
    }

    @Override
    public ListIterator<T> listIterator(int indice) {
        validarIndiceInsercao(indice);

        return new ListIterator<>() {
            private No<T> proximo = indice == tamanho ? null : obterNo(indice);
            private No<T> ultimoRetornado;
            private int cursor = indice;

            @Override
            public boolean hasNext() {
                return proximo != null;
            }

            @Override
            public T next() {
                if (proximo == null) throw new NoSuchElementException();
                ultimoRetornado = proximo;
                proximo = proximo.proximo;
                cursor++;
                return ultimoRetornado.valor;
            }

            @Override
            public boolean hasPrevious() {
                return proximo == null ? fim != null : proximo.anterior != null;
            }

            @Override
            public T previous() {
                No<T> anterior = proximo == null ? fim : proximo.anterior;
                if (anterior == null) throw new NoSuchElementException();
                proximo = anterior;
                ultimoRetornado = anterior;
                cursor--;
                return anterior.valor;
            }

            @Override
            public int nextIndex() {
                return cursor;
            }

            @Override
            public int previousIndex() {
                return cursor - 1;
            }

            @Override
            public void remove() {
                if (ultimoRetornado == null) throw new IllegalStateException();

                No<T> seguinte = ultimoRetornado.proximo;
                ListaDupla.this.desligar(ultimoRetornado);

                if (proximo == ultimoRetornado) {
                    proximo = seguinte;
                } else {
                    cursor--;
                }

                ultimoRetornado = null;
            }

            @Override
            public void set(T valor) {
                if (ultimoRetornado == null) throw new IllegalStateException();
                ultimoRetornado.valor = valor;
            }

            @Override
            public void add(T valor) {
                ListaDupla.this.add(cursor, valor);
                cursor++;
                ultimoRetornado = null;
            }
        };
    }

    private T desligar(No<T> no) {
        No<T> anterior = no.anterior;
        No<T> proximo = no.proximo;

        if (anterior == null) {
            inicio = proximo;
        } else {
            anterior.proximo = proximo;
        }

        if (proximo == null) {
            fim = anterior;
        } else {
            proximo.anterior = anterior;
        }

        T valor = no.valor;
        no.valor = null;
        no.anterior = null;
        no.proximo = null;

        tamanho--;
        modCount++;
        return valor;
    }

    private No<T> obterNo(int indice) {
        if (indice < tamanho / 2) {
            No<T> atual = inicio;
            for (int i = 0; i < indice; i++) {
                atual = atual.proximo;
            }
            return atual;
        }

        No<T> atual = fim;
        for (int i = tamanho - 1; i > indice; i--) {
            atual = atual.anterior;
        }
        return atual;
    }

    private void validarIndice(int indice) {
        if (indice < 0 || indice >= tamanho) {
            throw new IndexOutOfBoundsException(
                "Índice: " + indice + ", tamanho: " + tamanho
            );
        }
    }

    private void validarIndiceInsercao(int indice) {
        if (indice < 0 || indice > tamanho) {
            throw new IndexOutOfBoundsException(
                "Índice: " + indice + ", tamanho: " + tamanho
            );
        }
    }
}
