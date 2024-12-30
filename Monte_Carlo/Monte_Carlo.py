import xlwings as xw
import math
import random

def monte_carlo_option_pricing(S, K, T, r, sigma, num_simulations=100000, option_type="call"):
    '''
    Calculate the price of an option using the Monte Carlo method.

    Parameters:
    S (float): Current stock price
    K (float): Strike price
    T (float): Time to maturity in years
    r (float): Risk-free interest rate (annual, continuously compounded)
    sigma (float): Volatility of the stock's returns (annual standard deviation)
    num_simulations (int): Number of Monte Carlo simulations
    option_type (str): "call" for a call option, "put" for a put option

    Returns:
    price (float): Option price
    '''
    payoffs = []

    for _ in range(num_simulations):
        # Simulate the end-of-period stock price using the geometric Brownian motion model
        z = random.gauss(0, 1)  # Standard normal random variable
        ST = S * math.exp((r - 0.5 * sigma**2) * T + sigma * math.sqrt(T) * z)

        # Calculate the payoff for the option
        if option_type == "call":
            payoff = max(ST - K, 0)
        elif option_type == "put":
            payoff = max(K - ST, 0)
        else:
            raise ValueError("Invalid option type. Use 'call' or 'put'.")

        payoffs.append(payoff)

    # Discount the average payoff to present value
    price = math.exp(-r * T) * (sum(payoffs) / num_simulations)
    return price

if __name__ == "__main__":
    S = 100      # Current stock price
    K = 110      # Strike price
    T = 1        # Time to maturity (1 year)
    r = 0.05     # Risk-free rate (5%)
    sigma = 0.2  # Volatility (20%)
    num_simulations = 100000  # Number of Monte Carlo simulations

    call_price = monte_carlo_option_pricing(S, K, T, r, sigma, num_simulations, option_type="call")
    put_price = monte_carlo_option_pricing(S, K, T, r, sigma, num_simulations, option_type="put")

    print(f"Call Option Price (Monte Carlo): {call_price:.2f}")
    print(f"Put Option Price (Monte Carlo): {put_price:.2f}")
