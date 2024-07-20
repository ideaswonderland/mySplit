def ParaCevir(Para, PBirim="Lira", KBirim="Kuruş"):
    try:
        Para = float(Para)
    except ValueError:
        return "GİRİLEN DEĞER SAYI DEĞİL!"

    Lira, Kurus = "{:.2f}".format(abs(Para)).split(".")
    
    def Cevir(SayiStr):
        Birler = ["", "bir", "iki", "üç", "dört", "beş", "altı", "yedi", "sekiz", "dokuz"]
        Onlar = ["", "on", "yirmi", "otuz", "kırk", "elli", "altmış", "yetmiş", "seksen", "doksan"]
        Binler = ["trilyon", "milyar", "milyon", "bin", ""]

        SayiStr = SayiStr.zfill(15)
        Rakam = [int(c) for c in SayiStr]

        Sonuc = ""
        for i in range(0, 15, 3):
            c = Rakam[i:i+3]
            if c[0] == 1:
                e = "yüz" if c[1] != 0 or c[2] != 0 else "bir"
            elif c[0] != 0:
                e = Birler[c[0]] + "yüz"
            else:
                e = ""
            e += Onlar[c[1]] + Birler[c[2]]
            if e:
                e += Binler[i//3]
            if i == 9 and e == "birbin":
                e = "bin"
            Sonuc += e
        return Sonuc.capitalize()

    return ("Eksi " if Para < 0 else "") + Cevir(Lira) + " " + PBirim + (" " + Cevir(Kurus) + " " + KBirim if int(Kurus) else "")

