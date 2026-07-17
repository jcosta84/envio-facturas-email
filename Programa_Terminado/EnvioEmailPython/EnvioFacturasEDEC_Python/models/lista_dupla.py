from __future__ import annotations
from collections.abc import Iterable, Iterator
from typing import Generic, Optional, TypeVar, Callable

T = TypeVar("T")

class No(Generic[T]):
    __slots__ = ("valor", "anterior", "proximo")
    def __init__(self, valor: T):
        self.valor = valor
        self.anterior: Optional[No[T]] = None
        self.proximo: Optional[No[T]] = None

class ListaDupla(Generic[T]):
    def __init__(self, valores: Iterable[T] | None = None):
        self.inicio: Optional[No[T]] = None
        self.fim: Optional[No[T]] = None
        self._tamanho = 0
        if valores:
            for valor in valores: self.adicionar(valor)

    def __len__(self): return self._tamanho
    def __iter__(self) -> Iterator[T]:
        atual = self.inicio
        while atual:
            yield atual.valor
            atual = atual.proximo

    def adicionar(self, valor: T):
        no = No(valor)
        if self.fim is None: self.inicio = self.fim = no
        else:
            no.anterior = self.fim
            self.fim.proximo = no
            self.fim = no
        self._tamanho += 1

    inserir_fim = adicionar

    def inserir_inicio(self, valor: T):
        no = No(valor)
        if self.inicio is None: self.inicio = self.fim = no
        else:
            no.proximo = self.inicio
            self.inicio.anterior = no
            self.inicio = no
        self._tamanho += 1

    def pesquisar(self, criterio: Callable[[T], bool]) -> T | None:
        return next((v for v in self if criterio(v)), None)

    def filtrar(self, criterio: Callable[[T], bool]) -> "ListaDupla[T]":
        return ListaDupla(v for v in self if criterio(v))

    def remover(self, valor: T) -> bool:
        atual = self.inicio
        while atual:
            if atual.valor == valor:
                if atual.anterior: atual.anterior.proximo = atual.proximo
                else: self.inicio = atual.proximo
                if atual.proximo: atual.proximo.anterior = atual.anterior
                else: self.fim = atual.anterior
                self._tamanho -= 1
                return True
            atual = atual.proximo
        return False

    def limpar(self):
        self.inicio = self.fim = None
        self._tamanho = 0

    def para_lista(self) -> list[T]: return list(self)
