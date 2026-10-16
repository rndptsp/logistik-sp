#!/usr/bin/env python3
"""Encrypted inbox in the (public) repo: master files sealed to a build-only key, opened in GitHub Actions.

    repo_inbox.py seal <slot> <file.xlsx>   (anywhere) -> inbox_repo/<slot>-<time>.enc, sealed to tools/inbox_pub.pem
    repo_inbox.py open <home> <inbox>       (build)    -> <inbox>/<slot>.xlsx, same as an admin upload of that slot

The private key lives only in the private master store (R2: config/inbox_key.pem). The first build creates it and
publishes the public half as tools/inbox_pub.pem. A sealed file is RSA-OAEP(SHA-256) + AES-256-GCM: unreadable
without the store key, so it is safe in the public repo. An admin upload of the same slot wins over a sealed file.
Opened files are listed in <home>/_repo_inbox.txt; the build commit removes them. Only file names are printed.
"""
import base64, glob, json, os, sys, time
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PUB = os.path.join(REPO, 'tools', 'inbox_pub.pem')
DIR = os.path.join(REPO, 'inbox_repo')
SLOTS = ('rilis', 'so', 'tactical', 'snop', 'monitoring', 'gudang', 'kapal')
OAEP = padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None)
b64, unb64 = (lambda b: base64.b64encode(b).decode()), base64.b64decode


def seal(slot, path):
    if slot not in SLOTS: sys.exit('slot tidak dikenal: ' + slot)
    pub = serialization.load_pem_public_key(open(PUB, 'rb').read())
    key, iv = AESGCM.generate_key(bit_length=256), os.urandom(12)
    ct = AESGCM(key).encrypt(iv, open(path, 'rb').read(), slot.encode())
    os.makedirs(DIR, exist_ok=True)
    out = os.path.join(DIR, '%s-%s.enc' % (slot, time.strftime('%Y%m%d%H%M%S')))
    json.dump({'v': 1, 'slot': slot, 'ek': b64(pub.encrypt(key, OAEP)), 'iv': b64(iv), 'ct': b64(ct)}, open(out, 'w'))
    print('tersegel:', os.path.relpath(out, REPO))


def open_(home, inbox):
    kp = os.path.join(home, 'config', 'inbox_key.pem')
    if not os.path.exists(kp):   # first run: create the build-only key pair
        k = rsa.generate_private_key(public_exponent=65537, key_size=3072)
        os.makedirs(os.path.dirname(kp), exist_ok=True)
        open(kp, 'wb').write(k.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        open(PUB, 'wb').write(k.public_key().public_bytes(serialization.Encoding.PEM, serialization.PublicFormat.SubjectPublicKeyInfo))
        open(os.path.join(home, '_push_needed'), 'w').write('inbox key')
        print('repo-inbox: kunci baru dibuat (publik: tools/inbox_pub.pem)')
    priv = serialization.load_pem_private_key(open(kp, 'rb').read(), None)
    done = []
    for f in sorted(glob.glob(os.path.join(DIR, '*.enc'))):
        name = os.path.relpath(f, REPO)
        try:
            o = json.load(open(f)); slot = o['slot']
            if slot not in SLOTS: raise ValueError('slot')
            dest = os.path.join(inbox, slot + '.xlsx')
            if os.path.exists(dest):
                print('repo-inbox:', name, 'dilewati (slot', slot, 'sudah diupload lewat admin)'); continue
            data = AESGCM(priv.decrypt(unb64(o['ek']), OAEP)).decrypt(unb64(o['iv']), unb64(o['ct']), slot.encode())
            os.makedirs(inbox, exist_ok=True); open(dest, 'wb').write(data)
            done.append(name); print('repo-inbox:', name, '-> slot', slot)
        except Exception as e:
            print('repo-inbox: GAGAL membuka', name, '(' + type(e).__name__ + ')')
    open(os.path.join(home, '_repo_inbox.txt'), 'w').write('\n'.join(done))


if __name__ == '__main__':
    if len(sys.argv) == 4 and sys.argv[1] == 'seal': seal(sys.argv[2], sys.argv[3])
    elif len(sys.argv) == 4 and sys.argv[1] == 'open': open_(sys.argv[2], sys.argv[3])
    else: sys.exit(__doc__)
