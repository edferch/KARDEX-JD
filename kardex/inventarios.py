"""Selección de inventario activo (A / B / C) para la sesión actual."""
from flask import session

INVENTARIOS = ['A', 'B', 'C']


def obtener_inventario_actual():
    inv = session.get('inventario_actual', 'A')
    return inv if inv in INVENTARIOS else 'A'
