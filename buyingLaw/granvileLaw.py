def is_possible_4th_law(analysis_data):
    INTERVAL = 120
    is_success = False

    if analysis_data is None or len(analysis_data) < INTERVAL:
        is_success = False
    else:
        curr_moving_average_price = sum([int(i[1]) for i in analysis_data[:INTERVAL]]) / INTERVAL
        is_moving_average_price_in_today_bar = False
        curr_high_price = None
        prev_low_price = None
        high_price = analysis_data[0][6]
        low_price = analysis_data[0][7]

        if low_price <= curr_moving_average_price <= high_price:
            is_moving_average_price_in_today_bar = True
            curr_high_price = high_price

        if is_moving_average_price_in_today_bar is True:
            prev_moving_average_price = 0
            is_longer_than_minimum_period = False
            idx = 1

            while True:
                if len(analysis_data[idx:]) < INTERVAL:
                    break

                prev_moving_average_price = sum([int(i[1]) for i in analysis_data[idx:INTERVAL + idx]]) / INTERVAL

                _high_price = analysis_data[idx][6]
                _low_price = analysis_data[idx][7]

                if prev_moving_average_price <= _high_price and idx <= 20:
                    is_longer_than_minimum_period = False
                    break

                elif _low_price > prev_moving_average_price and idx > 20:
                    is_longer_than_minimum_period = True
                    prev_low_price = _low_price
                    break

                idx += 1

            if is_longer_than_minimum_period is True:
                if curr_moving_average_price > prev_moving_average_price and curr_high_price > prev_low_price:
                    is_success = True

    return is_success
