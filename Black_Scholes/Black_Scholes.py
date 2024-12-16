import xlwings as xw
import math

def cnd(x):
    '''
    Calculate the cumulative distribution function (CDF) of the standard normal distribution.
    Uses a numerical approximation for the error function (erf).

    Parameters:
    x (float): The value to evaluate the CDF at.

    Returns:
    float: CDF of the standard normal distribution at x.
    '''
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))

def black_scholes(S, K, T, r, sigma, option_type="call"):
    '''
    Calculate the Black-Scholes price of an option.

    Parameters:
    S (float): Current stock price
    K (float): Strike price
    T (float): Time to maturity in years
    r (float): Risk-free interest rate (annual, continuously compounded)
    sigma (float): Volatility of the stock's returns (annual standard deviation)
    option_type (str): "call" for a call option, "put" for a put option

    Returns:
    price (float): Option price
    '''

    d1 = (math.log(S / K) + (r + 0.5 * sigma**2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)

    if option_type == "call":
        price = S * cnd(d1) - K * math.exp(-r * T) * cnd(d2)

    elif option_type == "put":
        price = K * math.exp(-r * T) * cnd(-d2) - S * cnd(-d1)

    return price

if __name__ == "__main__":
    S = 100     # Current stock price
    K = 110     # Strike price
    T = 1       # Time to maturity (1 year)
    r = 0.05    # Risk-free rate (5%)
    sigma = 0.2 # Volatility (20%)

    call_price = black_scholes(S, K, T, r, sigma, option_type="call")
    put_price = black_scholes(S, K, T, r, sigma, option_type="put")

    print(f"Call Option Price: {call_price:.2f}")
    print(f"Put Option Price: {put_price:.2f}")
