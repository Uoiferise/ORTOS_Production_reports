from resources.resource_manager import ResourceManager


def main():
    data = ResourceManager.get_data(report_name='input_data/titanium_base/titanium_base_info.xlsx')
    test_name = '38709 ТО LM Bell (GEO) Implantium full G/H=1.3 H=5.3 с позиционером (арт. LL2-DER13-H) V.1 / / БЕЗ ВИНТА'

    for key, value in data[test_name].get_info().items():
        print(f'{key}: {value}')


if __name__ == '__main__':
    main()
