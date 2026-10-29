import base64
import json
import time

import pytest

from m365_mcp.links.tokens import LinkRefused, LinkSigner

K1, K2, E1, E2 = (bytes([i]) * 32 for i in (1, 2, 3, 4))
NOW = 1_800_000_000


def signer(sign=(K1,), enc=(E1,)):
    return LinkSigner(sign, enc)


def issue(s=None, **kw):
    args = dict(op="d", kind="attachment", target={"m": "MSG", "a": "ATT"}, person="p1", token="tok-abc",
                token_expires=NOW + 3600, ttl=300, now=NOW)
    args.update(kw)
    return (s or signer()).issue(**args)


def test_round_trip_carries_target_person_and_token():
    link = signer().verify(issue(), op="d", now=NOW + 10)
    assert link.target == {"m": "MSG", "a": "ATT"}
    assert link.person == "p1" and link.token == "tok-abc" and link.kind == "attachment"


def test_the_token_is_not_readable_in_the_link():
    value = issue(token="SECRET-BEARER")
    payload = base64.urlsafe_b64decode(value.split(".")[0] + "==")
    assert b"SECRET-BEARER" not in payload


@pytest.mark.parametrize("mangle", [
    lambda v: v[:-1] + ("A" if v[-1] != "A" else "B"),                     # signature byte
    lambda v: "x" + v,                                                     # payload byte
    lambda v: v.split(".")[0],                                             # no signature
    lambda v: "",                                                          # empty
    lambda v: "%%%.%%%",                                                   # not base64
    lambda v: "e30." + v.split(".")[1],                                    # other payload, old signature
])
def test_tampered_or_garbage_links_are_refused(mangle):
    with pytest.raises(LinkRefused):
        signer().verify(mangle(issue()), op="d", now=NOW + 10)


def test_payload_edited_and_resigned_without_the_key_is_refused():
    value = issue()
    body = json.loads(base64.urlsafe_b64decode(value.split(".")[0] + "=="))
    body["p"] = "someone-else"
    forged = base64.urlsafe_b64encode(json.dumps(body).encode()).rstrip(b"=").decode() + "." + value.split(".")[1]
    with pytest.raises(LinkRefused):
        signer().verify(forged, op="d", now=NOW + 10)


def test_wrong_operation_is_refused():
    with pytest.raises(LinkRefused):
        signer().verify(issue(), op="u", now=NOW + 10)


def test_expired_is_refused():
    with pytest.raises(LinkRefused):
        signer().verify(issue(), op="d", now=NOW + 301)


def test_expiry_never_outlives_the_token():
    link = signer().verify(issue(token_expires=NOW + 60, ttl=300), op="d", now=NOW + 1)
    assert link.expires == NOW + 60


def test_ttl_is_capped():
    link = signer().verify(issue(ttl=10_000), op="d", now=NOW + 1)
    assert link.expires == NOW + 900


def test_rotation_old_key_still_verifies():
    old = issue(signer(sign=(K1,), enc=(E1,)))
    rotated = signer(sign=(K2, K1), enc=(E2, E1))
    assert rotated.verify(old, op="d", now=NOW + 10).token == "tok-abc"


def test_removed_key_no_longer_verifies():
    old = issue(signer(sign=(K1,), enc=(E1,)))
    with pytest.raises(LinkRefused):
        signer(sign=(K2,), enc=(E1,)).verify(old, op="d", now=NOW + 10)


@pytest.mark.parametrize("sign, enc", [((), (E1,)), ((K1,), ()), ((b"short",), (E1,))])
def test_no_or_weak_key_refuses_to_sign(sign, enc):
    with pytest.raises(ValueError):
        LinkSigner(sign, enc)


def test_upload_kind_cannot_be_issued_as_a_download():
    with pytest.raises(ValueError):
        issue(op="d", kind="draft_body")


def test_upload_carries_limits():
    s = signer()
    v = s.issue(op="u", kind="draft_body", target={"d": "D"}, person="p1", token="t", token_expires=NOW + 3600,
                ttl=180, max_bytes=2048, content_type="text/html", now=NOW)
    link = s.verify(v, op="u", now=NOW + 5)
    assert link.max_bytes == 2048 and link.content_type == "text/html"


def test_with_a_store_a_link_is_a_short_id_that_verifies_like_the_signed_value():
    # the signed value runs to ~6,300 characters; with a store the link carries a
    # 22-character id and the value stays on the server.
    from m365_mcp.links.store import LinkStore
    sign, enc = b"s" * 32, b"e" * 32
    signer = LinkSigner((sign,), (enc,), LinkStore())
    long_token = "t" * 4000
    value = signer.issue(op="d", kind="drive", target={"i": "f1"}, person="p", token=long_token,
                         token_expires=int(time.time()) + 3600, ttl=300)
    assert len(value) < 40 and "." not in value
    link = signer.verify(value, op="d")
    assert link.token == long_token and link.target == {"i": "f1"}
    with pytest.raises(LinkRefused):
        signer.verify(value, op="u")  # the operation is still checked on the stored value
    with pytest.raises(LinkRefused):
        signer.verify(value[:-1] + ("A" if value[-1] != "A" else "B"), op="d")
    with pytest.raises(LinkRefused):
        signer.verify(value, op="d", now=time.time() + 301)
    with pytest.raises(LinkRefused):
        LinkSigner((sign,), (enc,), LinkStore()).verify(value, op="d")  # another process's id is unknown here
