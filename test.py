import scipy
from scipy.stats import norm


def FABI(results):
    footprint = (12) ** 2  # inches squared
    weight = 1.12  # lb
    design_life = 100  # years
    construction_cost = 2500000 * (weight ** 2) + 6 * (10 ** 6)
    land_cost = 35000 * footprint
    annual_building_cost = (land_cost + construction_cost) / design_life
    annual_revenue = 430300
    equipment_cost = 20000000
    return_period_1 = 50
    return_period_2 = 300
    max_disp = results[1]  # mm
    apeak_1 = results[0]  # g's
    xpeak_1 = 100 * max_disp / 1524  # % roof drift
    structural_damage_1 = scipy.stats.norm(1.5, 0.5).cdf(xpeak_1)
    equipment_damage_1 = scipy.stats.norm(1.75, 0.7).cdf(apeak_1)
    economic_loss_1 = structural_damage_1 * construction_cost + equipment_damage_1 * equipment_cost
    annual_economic_loss_1 = economic_loss_1 / return_period_1
    structural_damage_2 = 0.5
    equipment_damage_2 = 0.5
    economic_loss_2 = structural_damage_2 * construction_cost + equipment_damage_2 * equipment_cost
    annual_economic_loss_2 = economic_loss_2 / return_period_2
    annual_seismic_cost = annual_economic_loss_1 + annual_economic_loss_2
    fabi = annual_revenue - annual_building_cost - annual_seismic_cost
    return fabi

a = [0.5,5]
b = FABI(a)
print(b)
