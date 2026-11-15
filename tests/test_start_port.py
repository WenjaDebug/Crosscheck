"""Doppelt gestartet oder Port belegt: Gegenprobe sagt es in einem Satz, statt mit einer Fehlermeldung abzubrechen."""

import socket

from voiceagent import addin
from voiceagent.__main__ import _free, choose_port


def test_free_port_is_taken_as_it_is():
    assert choose_port(8765, busy=lambda p: False) == (8765, "")


def test_second_start_just_opens_the_browser():
    port, said = choose_port(8765, busy=lambda p: True, ours=lambda url: url == "http://127.0.0.1:8765")
    assert port is None and "läuft schon" in said


def test_other_program_on_the_port_gives_the_next_free_one_but_never_the_panel_port():
    taken = {8765, 8766}
    port, said = choose_port(8765, busy=lambda p: True, ours=lambda url: False, free=lambda p: p not in taken)
    assert port == 8767 and "Port 8765 hat schon ein anderes Programm" in said
    port, _ = choose_port(addin.PORT - 1, busy=lambda p: True, ours=lambda url: False, free=lambda p: True)
    assert port == addin.PORT + 1
    port, said = choose_port(8765, busy=lambda p: True, ours=lambda url: False, free=lambda p: False)
    assert port == 8765 and "keinen freien" in said


def test_free_sees_a_real_listener():
    with socket.socket() as server:
        server.bind(("127.0.0.1", 0))
        server.listen()
        assert _free(server.getsockname()[1]) is False
