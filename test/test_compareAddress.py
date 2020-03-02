

def test_compareAdress():
    addr1 = {'nation': '中国', 'province': '福建省', 'city': '漳州市', 'district': '芗城区', 'street': '延安北路', 'street_number': '延安北路52号'}
    addr2 = {'nation': '中国', 'province': '湖北省', 'city': '黄冈市', 'district': '黄梅县', 'street': '113县道', 'street_number': '113县道'}
    addr3 = {'nation': '中国', 'province': '江西省', 'city': '赣州市', 'district': '信丰县', 'street': '南山西路', 'street_number': '南山西路'}
    addr4 = {'nation': '中国', 'province': '福建省', 'city': '福州市', 'district': '闽侯县', 'street': '科技东路', 'street_number': '科技东路'}
    addr5 = {'nation': '中国', 'province': '福建省', 'city': '福州市', 'district': '长乐区', 'street': '', 'street_number': ''}
    addr6 = {'nation': '中国', 'province': '福建省', 'city': '漳州市', 'district': '云霄县', 'street': '577县道', 'street_number': '577县道'}

    print(compareAdress(addr1, addr1))
    print(compareAdress(addr1, addr2))
    print(compareAdress(addr1, addr3))
    print(compareAdress(addr1, addr4))
    print(compareAdress(addr1, addr5))
    print(compareAdress(addr1, addr6))
    print(compareAdress(addr2, addr1))
    print(compareAdress(addr2, addr2))
    print(compareAdress(addr2, addr3))
    print(compareAdress(addr2, addr4))
    print(compareAdress(addr2, addr5))
    print(compareAdress(addr2, addr6))
    print(compareAdress(addr3, addr1))
    print(compareAdress(addr3, addr2))
    print(compareAdress(addr3, addr4))
    print(compareAdress(addr3, addr5))
    print(compareAdress(addr3, addr6))
    print(compareAdress(addr4, addr1))
    print(compareAdress(addr4, addr2))
    print(compareAdress(addr4, addr3))
    print(compareAdress(addr4, addr4))
    print(compareAdress(addr4, addr5))
    print(compareAdress(addr4, addr6))
    print(compareAdress(addr5, addr1))
    print(compareAdress(addr5, addr2))
    print(compareAdress(addr5, addr3))
    print(compareAdress(addr5, addr4))
    print(compareAdress(addr5, addr5))

    
if __name__ == '__main__':

    test_compareAdress()