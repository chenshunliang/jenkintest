from flask import render_template, redirect, url_for, abort, flash, request, current_app, make_response, session, Response, send_file, jsonify
from flask.ext.login import login_required, current_user
from . import calculator

from .. import db

import pandas as pd
import datetime
import json
import time
import numpy as np
import random
import os
import string

# set pandas header format
# pd.core.format.header_style = None


###################--Mission Statement ###########################
#Goal: This module is to build up a calculator to figure out the #
#exact box office to make investors break even. And we would also#
# calculate what ROI would we make under other circumstances.   #
# 5 pm, Oct 16th, 2016                             #
#############################-end-################################


####################--History review--###########################
# version evolution
"""
In the initial version called calculatorBeta, actually we make all the computations included in the html javascript file. In that way, we would make it no related to the server, which theoretically could lessen the server's pressure. However, in the later version called calculatorModern, we instead make all the computations back on the server because we believe it would much easier to manipulate ms excel files using python rather than javascript. After all, we are much more familier with python.
"""

# module structure
"""
#Updated at about 17:30 on Oct 16th,2016
Compared to calculatorModern, we make serveral changes and we would present and analyse these changes one by one and try to explain why we would do like that.

#1 change: we add ladder to allow different parameters for different box office ranges.
Well, this is quite necessary change as it is not uncommon that the distributing department would negotiate with partners in this way. And actually it is.

#2 change: we split one break-even analysis into three.
Generally, there may be two roles for wepiao in investment activities, distributor and investor. Given wepiao probably would participate both distribution and investment activities, it would be quite necessary for us to present when wepiao would break even when acting as investor or distributor, respectively and both. Therefore, we have to output three reports.
"""
#############################-end-################################


# calculatorModern
@calculator.route('/calculatorModern', methods=['GET', 'POST'])
def calculatorModern():
    """
    This function is to provide our users with a platform to calculate rate of return when investing a movie.
    """
    return(render_template('calculator/calculator.html'))
# end of calculatorModern


# boxoffice_high
def Boxoffice_high(parameter_list, wepiao=True):
    """
    Actually in the next section, we find that it's too complicated for set conditions for boxoffice_high and boxoffice_low. Strictly speaking, we can set all boxoffice_low to be zero. However, to deal with boxoffice_high, things get complicated as we want to enumerate all the circumstances.
    """
    # case 0: check copyrightIncome
    # if copyrightIncome already covers cost, there is no need to proceed, we
    # just have to stop
    if wepiao == False and parameter_list['copyrightIncome'] >= parameter_list['production']:
        boxoffice_high = 0
    elif wepiao == True and parameter_list['copyrightIncome_wepiao'] >= parameter_list['production_wepiao'] + parameter_list['propaganda_wepiao']:
        boxoffice_high = 0
    # end of case 0
    else:
        # case 1: zzbPercentage or taxPercentage to be 100% or boxPercentage to be 0%
        #    for film, wepiao = False
        # we just need to compare production cost with copyrightIncome. If
        # latter is not less than the former, there is no hope. Otherwise, we
        # set boxoffice_high to be zero directly. What is tricky here is that
        # if copyrightIncome is equal with production cost, we also set
        # copyrightIncome to be zero. But as I know later, this comparison is
        # meaningful only if we want to know whether the film can break even,
        # otherwise we just need to set boxoffice_high to be zero no matter
        # which one is bigger. It is the same to wepiao.
        if (parameter_list['zzbPercentage'] == 100 or parameter_list['taxPercentage'] == 100 or parameter_list['boxPercentage'] == 0):
            boxoffice_high = 0
        # end of case 1
        # case 2: proxyPercentage is 100%
        elif parameter_list['proxyPercentage'] == 100:
            # If proxyPercentage is 100, clearly, for film, this case has the
            # same effect on film's performance as the above paramaters
            if wepiao == False:
                boxoffice_high = 0
            else:
                # If proxyPercentage is 100, for wepiao, however, things get
                # complicated. On the one hand, there is no possibility that
                # wepiao could get boxoffice by propaganda or production; On
                # ther other hand, if wepiao could receive return by means of
                # proxy, this is no big problem.
                if parameter_list['proxyPercentage_wepiao'] == 0:
                    boxoffice_high = 0
                else:
                    boxoffice_high = 100 * (parameter_list['production_wepiao'] + parameter_list['propaganda_wepiao']) / ((1 - parameter_list['zzbPercentage'] / 100 - parameter_list[
                        'taxPercentage'] / 100) * (parameter_list['boxPercentage'] / 100) * (parameter_list['proxyPercentage_wepiao'] / 100))
        # end of case 2
        # case 3: remaining regular cases for film
        else:
            if wepiao == False:
                boxoffice_high = 100 * (parameter_list['production'] + parameter_list['propaganda']) / ((1 - parameter_list['zzbPercentage'] / 100 - parameter_list[
                    'taxPercentage'] / 100) * (parameter_list['boxPercentage'] / 100) * (1 - parameter_list['proxyPercentage'] / 100))
            # end of case 3
            else:
                # case 4: if there is premium for production_wepiao or
                # propaganda_wepiao. We check if proxyPercentage or
                # productionPercentage_wepiao, if they are both zero, there is
                # no hope, otherwise there would be great hope.
                if parameter_list['proxyPercentage_wepiao'] == 0 and parameter_list['productionPercentage_wepiao'] == 0 and parameter_list['propaganda'] * parameter_list['propagandaPercentage_wepiao'] / 100 + parameter_list['copyrightIncome_wepiao'] < parameter_list['propaganda_wepiao']:
                    # under this case, there is no way to break even because
                    # the maximum propaganda wepiao can get is less than what
                    # required by customized propaganda cost
                    boxoffice_high = 0
                elif parameter_list['proxyPercentage_wepiao'] != 0 and parameter_list['productionPercentage_wepiao'] == 0:
                    boxoffice_high = 100 * (parameter_list['production_wepiao'] + parameter_list['propaganda_wepiao']) / ((1 - parameter_list['zzbPercentage'] / 100 - parameter_list[
                        'taxPercentage'] / 100) * (parameter_list['boxPercentage'] / 100) * (parameter_list['proxyPercentage_wepiao'] / 100))
                elif parameter_list['proxyPercentage_wepiao'] == 0 and parameter_list['productionPercentage_wepiao'] != 0:
                    boxoffice_high = 100 * (parameter_list['production_wepiao'] + parameter_list['propaganda_wepiao']) / ((1 - parameter_list['zzbPercentage'] / 100 - parameter_list[
                        'taxPercentage'] / 100) * (parameter_list['boxPercentage'] / 100) * (1 - parameter_list['proxyPercentage'] / 100) * (parameter_list['productionPercentage_wepiao'] / 100))
                else:
                    boxoffice_high = 100 * (parameter_list['production_wepiao'] + parameter_list['propaganda_wepiao']) / ((1 - parameter_list['zzbPercentage'] / 100 - parameter_list[
                        'taxPercentage'] / 100) * (parameter_list['boxPercentage'] / 100) * (1 - parameter_list['proxyPercentage'] / 100) * (parameter_list['propagandaPercentage_wepiao'] / 100))
                # end of case 4
    return boxoffice_high
# end of Boxoffice_high


# calculate income item by item
# issue income
def GetIssueIncome(boxoffice, boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage, result):
    """
    Input boxoffice and boxPercentage ladder, we output issue income.
    If boxofficePercentage_ladder's length is less than three, just use boxPercentage from parameter_list.
    """
    if len(boxofficePercentage_ladder) < 3:
        issue_income = boxoffice * boxoffice_net_percetage * \
            parameter_list['boxPercentage'] / 100
    else:
        boxofficeRange_list = []

        i = 0
        j = 0
        range_max = int((len(boxofficePercentage_ladder) + 1) / 3)
        while j < range_max:
            if i == 0:
                boxofficeRange_list.append({'lower': boxofficePercentage_ladder['input_highest_boxPercentage_ladder_field_' + str(i)], 'higher': boxofficePercentage_ladder[
                                           'input_highest_boxPercentage_ladder_field_' + str(i)], 'value': boxofficePercentage_ladder['input_interested_boxPercentage_ladder_field_' + str(i)]})
                j += 1
            else:
                try:
                    boxofficeRange_list.append({'lower': boxofficePercentage_ladder['input_low_boxPercentage_ladder_field_' + str(i)], 'higher': boxofficePercentage_ladder[
                                               'input_high_boxPercentage_ladder_field_' + str(i)], 'value': boxofficePercentage_ladder['input_interested_boxPercentage_ladder_field_' + str(i)]})
                    j += 1
                except Exception as e:
                    i += 1
                    continue
            i += 1

        # sort the list of dict
        boxofficeRange_list_sorted = sorted(
            boxofficeRange_list, key=lambda user: user['lower'])

        # initialize issue_income
        issue_income = 0

        # calculate issue income range by range
        for range_i in boxofficeRange_list_sorted:
            if boxoffice < range_i['lower']:
                break
            elif boxoffice < range_i['higher'] or range_i['higher'] == range_i['lower']:
                issue_income += (boxoffice - range_i['lower']) * \
                    boxoffice_net_percetage * range_i['value'] / 100
            else:
                issue_income += (range_i['higher'] - range_i['lower']) * \
                    boxoffice_net_percetage * range_i['value'] / 100

            # if boxoffice == 10000:
            #    print(range_i['lower'], range_i['higher'], issue_income, boxoffice)

    # now we subtract despositCost
    depositCost = GetDepositCost(boxoffice, result[
                                 'depositCost_ladder'], parameter_list, copyrightIncome_key='depositCost_ladder', no_ladder_key='depositCost')

    issue_income = issue_income - depositCost

    return issue_income
# end of GetIssueIncome


# calculate income item by item
# proxy income, both for film and wepiao
def GetProxyIncome(boxoffice, boxofficePercentage_ladder, proxyPercentage_ladder, parameter_list, boxoffice_net_percetage, result, no_ladder_key='proxyPercentage'):
    """
    Input boxoffice and proxyPercentage_ladderr, we output proxy income.
    If proxyPercentage_ladder's length is less than three, just use proxyPercentage or proxyPercentage_wepiao from parameter_list.
    """
    if len(proxyPercentage_ladder) < 3:
        proxy_income = GetIssueIncome(boxoffice, boxofficePercentage_ladder, parameter_list,
                                      boxoffice_net_percetage, result) * parameter_list[no_ladder_key] / 100
    else:
        boxofficeRange_list = []

        i = 0
        j = 0
        range_max = int((len(proxyPercentage_ladder) + 1) / 3)
        while j < range_max:
            if i == 0:
                boxofficeRange_list.append({'lower': proxyPercentage_ladder['input_highest_proxyPercentage_ladder_field_' + str(i)], 'higher': proxyPercentage_ladder[
                                           'input_highest_proxyPercentage_ladder_field_' + str(i)], 'value': proxyPercentage_ladder['input_interested_proxyPercentage_ladder_field_' + str(i)]})
                j += 1
            else:
                try:
                    boxofficeRange_list.append({'lower': proxyPercentage_ladder['input_low_proxyPercentage_ladder_field_' + str(i)], 'higher': proxyPercentage_ladder[
                                               'input_high_proxyPercentage_ladder_field_' + str(i)], 'value': proxyPercentage_ladder['input_interested_proxyPercentage_ladder_field_' + str(i)]})
                    j += 1
                except Exception as e:
                    print(e)
                    i += 1
                    continue
            i += 1

        # sort the list of dict
        boxofficeRange_list_sorted = sorted(
            boxofficeRange_list, key=lambda user: user['lower'])

        # initialize issue_income
        proxy_income = 0

        # calculate issue income range by range, (boxoffice,
        # boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage,
        # result)
        for range_i in boxofficeRange_list_sorted:
            if boxoffice < range_i['lower']:
                break
            elif boxoffice < range_i['higher'] or range_i['higher'] == range_i['lower']:
                proxy_income += (max(0, GetIssueIncome(boxoffice, boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage, result)) - max(
                    0, GetIssueIncome(range_i['lower'], boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage, result))) * range_i['value'] / 100
            else:
                proxy_income += (max(0, GetIssueIncome(range_i['higher'], boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage, result)) - max(
                    0, GetIssueIncome(range_i['lower'], boxofficePercentage_ladder, parameter_list, boxoffice_net_percetage, result))) * range_i['value'] / 100
                # obviously we have to make sure that issue income must be
                # bigger than zero when calculating proxy income, this is a bug
                # hard to debug when there is a ladder for proxyPercentage

    proxy_income = max(0, proxy_income)

    return proxy_income
# end of GetProxyIncome


# calculate income item by item
# copyrightIncome, for film and wepiao respectively
def GetCopyrightIncome(boxoffice, copyrightIncome_ladder, parameter_list, copyrightIncome_key='copyrightIncome_ladder', no_ladder_key='copyrightIncome'):
    """
    Input boxoffice and proxyPercentage_ladderr, we output proxy income.
    If proxyPercentage_ladder's length is less than three, just use proxyPercentage or proxyPercentage_wepiao from parameter_list.
    """
    if len(copyrightIncome_ladder) < 3:
        if no_ladder_key == 'copyrightIncome':
            copyright_income = parameter_list[
                no_ladder_key] * (1 - parameter_list['copyrightProxyPercentage'] / 100)
        else:
            copyright_income = parameter_list[no_ladder_key]
    else:
        boxofficeRange_list = []

        i = 0
        j = 0
        range_max = int((len(copyrightIncome_ladder) + 1) / 3)
        while j < range_max:
            if i == 0:
                boxofficeRange_list.append({'lower': copyrightIncome_ladder['input_highest_' + copyrightIncome_key + '_field_' + str(i)], 'higher': copyrightIncome_ladder[
                                           'input_highest_' + copyrightIncome_key + '_field_' + str(i)], 'value': copyrightIncome_ladder['input_interested_' + copyrightIncome_key + '_field_' + str(i)]})
                j += 1
            else:
                try:
                    boxofficeRange_list.append({'lower': copyrightIncome_ladder['input_low_' + copyrightIncome_key + '_field_' + str(i)], 'higher': copyrightIncome_ladder[
                                               'input_high_' + copyrightIncome_key + '_field_' + str(i)], 'value': copyrightIncome_ladder['input_interested_' + copyrightIncome_key + '_field_' + str(i)]})
                    j += 1
                except Exception as e:
                    print(e)
                    i += 1
                    continue
            i += 1

        # sort the list of dict
        boxofficeRange_list_sorted = sorted(
            boxofficeRange_list, key=lambda user: user['lower'])

        # initialize issue_income
        copyright_income = 0

        # calculate issue income range by range
        for range_i in boxofficeRange_list_sorted:
            if boxoffice < range_i['lower']:
                break
            elif boxoffice < range_i['higher'] or range_i['higher'] == range_i['lower']:
                copyright_income = range_i['value']
            else:
                copyright_income = range_i['value']

    return copyright_income
# end of GetCopyrightIncome


# calculate income item by item
# depositCost, for film and wepiao respectively
def GetDepositCost(boxoffice, copyrightIncome_ladder, parameter_list, copyrightIncome_key='depositCost_ladder', no_ladder_key='depositCost'):
    """
    Input boxoffice and proxyPercentage_ladderr, we output proxy income.
    """
    if len(copyrightIncome_ladder) < 3:
        copyright_income = parameter_list[no_ladder_key]
    else:
        boxofficeRange_list = []

        i = 0
        j = 0
        range_max = int((len(copyrightIncome_ladder) + 1) / 3)
        while j < range_max:
            if i == 0:
                boxofficeRange_list.append({'lower': copyrightIncome_ladder['input_highest_' + copyrightIncome_key + '_field_' + str(i)], 'higher': copyrightIncome_ladder[
                                           'input_highest_' + copyrightIncome_key + '_field_' + str(i)], 'value': copyrightIncome_ladder['input_interested_' + copyrightIncome_key + '_field_' + str(i)]})
                j += 1
            else:
                try:
                    boxofficeRange_list.append({'lower': copyrightIncome_ladder['input_low_' + copyrightIncome_key + '_field_' + str(i)], 'higher': copyrightIncome_ladder[
                                               'input_high_' + copyrightIncome_key + '_field_' + str(i)], 'value': copyrightIncome_ladder['input_interested_' + copyrightIncome_key + '_field_' + str(i)]})
                    j += 1
                except Exception as e:
                    print(e)
                    i += 1
                    continue
            i += 1

        # sort the list of dict
        boxofficeRange_list_sorted = sorted(
            boxofficeRange_list, key=lambda user: user['lower'])

        # initialize issue_income
        copyright_income = 0

        # calculate issue income range by range
        for range_i in boxofficeRange_list_sorted:
            if boxoffice < range_i['lower']:
                break
            elif boxoffice < range_i['higher'] or range_i['higher'] == range_i['lower']:
                copyright_income += range_i['value']
            else:
                copyright_income += range_i['value']

    return copyright_income
# end of GetDepositCost


# return various income from a dozen inputs
def GetVriousItem(boxoffice_wepiao, parameter_list, boxoffice_net_percetage, result, even_percentage=0):
    """
    Write all various income into a dictionary to call.
    """
    # firstly, we calculate issue_income
    boxoffice_net = boxoffice_wepiao * boxoffice_net_percetage
    issue_income = GetIssueIncome(boxoffice_wepiao, result[
                                  'boxPercentage_ladder'], parameter_list, boxoffice_net_percetage, result)
    depositCost = GetDepositCost(boxoffice_wepiao, result[
                                 'depositCost_ladder'], parameter_list, copyrightIncome_key='depositCost_ladder', no_ladder_key='depositCost')
    proxy_income = GetProxyIncome(boxoffice_wepiao, result['boxPercentage_ladder'], result[
                                  'proxyPercentage_ladder'], parameter_list, boxoffice_net_percetage, result, no_ladder_key='proxyPercentage')
    net_issue_income = issue_income - proxy_income

    ################---film as a whole--######################################
    profit_film = net_issue_income - \
        parameter_list['propaganda'] - parameter_list['production']
    casts_profit = max(
        0, profit_film * parameter_list['castsPercentage'] / 100)
    net_profit_film = profit_film - casts_profit
    copyrightIncome = GetCopyrightIncome(boxoffice_wepiao, result[
                                         'copyrightIncome_ladder'], parameter_list, copyrightIncome_key='copyrightIncome_ladder', no_ladder_key='copyrightIncome')
    total_income_investor = copyrightIncome + \
        max(0, net_issue_income - parameter_list['propaganda'])
    total_profit_investor = total_income_investor - \
        parameter_list['production']
    try:
        roi_film_investor = 100 * total_profit_investor / \
            parameter_list['production']
    except:
        roi_film_investor = 'N/A'

    ###############--wepiao income--##########################################
    # therefore, we get #1 income: proxy_income_wepiao
    proxy_income_wepiao = GetProxyIncome(boxoffice_wepiao, result['boxPercentage_ladder'], result[
                                         'proxyPercentage_wepiao_ladder'], parameter_list, boxoffice_net_percetage, result, no_ladder_key='proxyPercentage_wepiao')

    # now we calculate #2 income: propaganda_income_wepiao
    propaganda_income_wepiao = min(parameter_list['propaganda'] * parameter_list['propagandaPercentage_wepiao'] / 100, max(
        0, net_issue_income * parameter_list['propagandaPercentage_wepiao'] / 100))

    # now #3 income: production_income_wepiao
    production_income_film = max(
        0, net_issue_income - parameter_list['propaganda'])
    if production_income_film < parameter_list['production']:
        production_income_wepiao = production_income_film * \
            parameter_list['productionPercentage_wepiao'] / 100
    else:
        production_income_wepiao = (parameter_list['production'] + (production_income_film - parameter_list['production']) * (
            1 - parameter_list['castsPercentage'] / 100)) * parameter_list['productionPercentage_wepiao'] / 100

    # 4 income: copyrightIncome_wepiao
    copyrightIncome_wepiao = GetCopyrightIncome(boxoffice_wepiao, result[
                                                'copyrightIncome_wepiao_ladder'], parameter_list, copyrightIncome_key='copyrightIncome_wepiao_ladder', no_ladder_key='copyrightIncome_wepiao')

    # 1 cost: propaganda_cost_wepiao
    propaganda_cost_wepiao = parameter_list['propaganda_wepiao']

    # 2 cost: production_cost_wepiao
    production_cost_wepiao = parameter_list['production_wepiao']

    # here we get total_income_wepiao
    total_income_wepiao_combined = proxy_income_wepiao + \
        propaganda_income_wepiao + production_income_wepiao + copyrightIncome_wepiao
    total_income_wepiao_distributor = proxy_income_wepiao + propaganda_income_wepiao
    total_income_wepiao_investor = production_income_wepiao + copyrightIncome_wepiao

    ###############--wepiao cost--############################################
    # 3 cost: depositCost_wepiao
    #depositCost_wepiao = GetDepositCost(boxoffice_wepiao, result['copyrightIncome_ladder'], parameter_list, copyrightIncome_key = 'depositCost_wepiao_ladder', no_ladder_key = 'depositCost_wepiao')

    # here, we get total_cost_wepiao
    total_cost_wepiao_combined = propaganda_cost_wepiao + production_cost_wepiao
    total_cost_wepiao_distributor = propaganda_cost_wepiao
    total_cost_wepiao_investor = production_cost_wepiao

    ###############--profit--#################################################
    profit_wepiao_combined = total_income_wepiao_combined - \
        total_cost_wepiao_combined * (1 + even_percentage / 100)
    profit_wepiao_distributor = total_income_wepiao_distributor - \
        total_cost_wepiao_distributor * (1 + even_percentage / 100)
    profit_wepiao_investor = total_income_wepiao_investor - \
        total_cost_wepiao_investor * (1 + even_percentage / 100)

    output = {'boxoffice_wepiao': boxoffice_wepiao, 'boxoffice_net': boxoffice_net, 'issue_income': issue_income, 'depositCost': depositCost, 'proxy_income': proxy_income, 'net_issue_income': net_issue_income, 'propaganda': parameter_list['propaganda'], 'production': parameter_list['production'], 'profit_film': profit_film, 'casts_profit': casts_profit, 'net_profit_film': net_profit_film, 'copyrightIncome': copyrightIncome, 'total_income_investor': total_income_investor, 'total_profit_investor': total_profit_investor, 'roi_film_investor': roi_film_investor, 'proxy_income_wepiao': proxy_income_wepiao, 'propaganda_income_wepiao': propaganda_income_wepiao, 'production_income_wepiao': production_income_wepiao,
              'copyrightIncome_wepiao': copyrightIncome_wepiao, 'total_income_wepiao_combined': total_income_wepiao_combined, 'total_income_wepiao_distributor': total_income_wepiao_distributor, 'total_income_wepiao_investor': total_income_wepiao_investor, 'propaganda_cost_wepiao': propaganda_cost_wepiao, 'production_cost_wepiao': production_cost_wepiao, 'total_cost_wepiao_combined': total_cost_wepiao_combined, 'total_cost_wepiao_distributor': total_cost_wepiao_distributor, 'total_cost_wepiao_investor': total_cost_wepiao_investor, 'profit_wepiao_combined': profit_wepiao_combined, 'profit_wepiao_distributor': profit_wepiao_distributor, 'profit_wepiao_investor': profit_wepiao_investor}
    # print(output)
    # print('\n')

    return output
# end of GetVriousItem


# calculate profit for wepiao as investor and distributor combined
def Profit_combined(boxoffice_high, preset_error, parameter_list, prescribed_loop_number, boxoffice_net_percetage, result, option='combined', even_percentage=0):
    """
    Calculate profit for wepiao.
    We believe that cost like depositCost should be put in before proxy income.
    """
    i = 0
    profit_wepiao = 0
    boxoffice_wepiao = 0

    while i == 0 or (profit_wepiao > preset_error or profit_wepiao < (-1) * preset_error):
        if profit_wepiao > preset_error:
            boxoffice_high = boxoffice_wepiao - 0.1 * preset_error
        else:
            boxoffice_low = boxoffice_wepiao + 0.1 * preset_error

        boxoffice_wepiao = 0.5 * (boxoffice_high + boxoffice_low)
        #boxoffice_net = boxoffice_wepiao * boxoffice_net_percetage

        # get various item from GetVriousItem(boxoffice_wepiao, parameter_list,
        # boxoffice_net_percetage, result)
        variousItems = GetVriousItem(boxoffice_wepiao, parameter_list,
                                     boxoffice_net_percetage, result, even_percentage=even_percentage)

        if option == 'combined':
            profit_wepiao = variousItems['profit_wepiao_combined']
        elif option == 'distributor':
            profit_wepiao = variousItems['profit_wepiao_distributor']
        else:
            profit_wepiao = variousItems['profit_wepiao_investor']

        # now deal with profit is always bigger than zero
        if boxoffice_wepiao < 100 * preset_error and boxoffice_wepiao > (-100) * preset_error:
            boxoffice_wepiao = 0
            break

        # if boxoffice is too high, while distributor profit is zero, we have
        # to cut boxoffice_high
        if option == 'distributor' and variousItems['net_issue_income'] > variousItems['propaganda'] + 100 * preset_error and even_percentage == 0:
            profit_wepiao = 1

        if option == 'combined' and variousItems['production_cost_wepiao'] == 0 and variousItems['net_issue_income'] > variousItems['propaganda'] + 100 * preset_error and even_percentage == 0:
            profit_wepiao = 1

        if option == 'investor' and variousItems['production_cost_wepiao'] == 0 and even_percentage == 0:
            boxoffice_wepiao = 0
            break

        # deal with special circumstances
        if i > prescribed_loop_number:
            break

        if i < 100:
            print(i)
            print(option + ': profit: {:.4f}, boxoffice_high: {:.4f}, boxoffice_low: {:.4f}, boxoffice_wepiao: {:.4f}'.format(
                profit_wepiao, boxoffice_high, boxoffice_low, boxoffice_wepiao))
            print(variousItems)
            print('\n')

        i += 1

    return boxoffice_wepiao
# end of Profit_combined


# calculate break-even boxoffice for film and wepiao, respectively
def BreakEvenBoxoffice(parameter_list, result, prescribed_loop_number=10000, even_percentage=0):
    """
    By passing various necessary parameters, we return break-even boxoffice for film and wepiao, respectively.
    Actually, an important issue here is to deal with those circumstances under which we cannot figure out the break-even boxoffice.
    For instance, given certain conditions, no matter how big the box office is, we may find out that the film investors would always suffer loss. In order to deal with these cases, we have to build some default values for break-even boxoffice. 
    Updated on Oct 12th, 2016: 
    We tear this big function into above three functions, please check.
    """
    preset_error = 0.000001 * min(1, parameter_list['production_wepiao'], parameter_list[
                                  'production'], parameter_list['propaganda_wepiao'])
    boxoffice_net_percetage = (
        1 - parameter_list['taxPercentage'] / 100 - parameter_list['zzbPercentage'] / 100)
    boxoffice_high = Boxoffice_high(parameter_list, wepiao=True)

    boxoffice_combined = Profit_combined(boxoffice_high, preset_error, parameter_list, prescribed_loop_number,
                                         boxoffice_net_percetage, result, option='combined', even_percentage=even_percentage)
    boxoffice_distributor = Profit_combined(boxoffice_high, preset_error, parameter_list, prescribed_loop_number,
                                            boxoffice_net_percetage, result, option='distributor', even_percentage=even_percentage)
    boxoffice_investor = Profit_combined(Boxoffice_high(parameter_list, wepiao=False), preset_error, parameter_list,
                                         prescribed_loop_number, boxoffice_net_percetage, result, option='investor', even_percentage=even_percentage)

    return boxoffice_distributor, boxoffice_investor, boxoffice_combined
# end of BreakEvenBoxoffice


# get return on calculator
def calculatorDataRow(boxoffice_start, parameter_list, boxoffice_net_percetage, result):
    """
    Get one row data for the data frame.
    """
    # calculate various item
    variousItems = GetVriousItem(
        boxoffice_start, parameter_list, boxoffice_net_percetage, result)

    # film as a whole
    boxoffice_net = variousItems['boxoffice_net']
    depositCost = variousItems['depositCost']
    issue_income = variousItems['issue_income']
    proxy_income = variousItems['proxy_income']
    net_issue_income = variousItems['net_issue_income']
    propaganda = variousItems['propaganda']
    production = variousItems['production']
    profit_film = variousItems['profit_film']
    casts_profit = variousItems['casts_profit']
    net_profit_film = variousItems['net_profit_film']
    copyrightIncome = variousItems['copyrightIncome']
    total_income_investor = variousItems['total_income_investor']
    total_profit_investor = variousItems['total_profit_investor']
    roi_film_investor = variousItems['roi_film_investor']

    # combined
    proxy_income_wepiao = variousItems['proxy_income_wepiao']
    propaganda_income_wepiao = variousItems['propaganda_income_wepiao']
    production_income_wepiao = variousItems['production_income_wepiao']
    copyrightIncome_wepiao = variousItems['copyrightIncome_wepiao']
    total_income_wepiao_combined = variousItems['total_income_wepiao_combined']
    propaganda_cost_wepiao = variousItems['propaganda_cost_wepiao']
    production_cost_wepiao = variousItems['production_cost_wepiao']
    total_cost_wepiao_combined = variousItems['total_cost_wepiao_combined']
    profit_wepiao_combined = variousItems['profit_wepiao_combined']

    # distributor
    total_income_wepiao_distributor = variousItems[
        'total_income_wepiao_distributor']
    total_cost_wepiao_distributor = variousItems[
        'total_cost_wepiao_distributor']
    profit_wepiao_distributor = variousItems['profit_wepiao_distributor']

    # investor
    total_income_wepiao_investor = variousItems['total_income_wepiao_investor']
    total_cost_wepiao_investor = variousItems['total_cost_wepiao_investor']
    profit_wepiao_investor = variousItems['profit_wepiao_investor']

    # we have to deal with what we know as zero cost
    # investor
    try:
        investor_roi = profit_wepiao_investor * 100 / total_cost_wepiao_investor
    except:
        investor_roi = 'N/A'

    # distirbutor
    try:
        distirbutor_roi = profit_wepiao_distributor * 100 / total_cost_wepiao_distributor
    except:
        distirbutor_roi = 'N/A'

    # combined roi
    try:
        combined_roi = profit_wepiao_combined * 100 / total_cost_wepiao_combined
    except:
        combined_roi = 'N/A'

    return [boxoffice_start, boxoffice_net, depositCost, issue_income, proxy_income, net_issue_income, propaganda, production, profit_film, casts_profit, net_profit_film, copyrightIncome, total_income_investor, total_profit_investor, roi_film_investor, '', proxy_income_wepiao, propaganda_income_wepiao, production_income_wepiao, copyrightIncome_wepiao, total_income_wepiao_combined, propaganda_cost_wepiao, production_cost_wepiao, total_cost_wepiao_combined, profit_wepiao_combined, combined_roi, '', proxy_income_wepiao, propaganda_income_wepiao, total_income_wepiao_distributor, total_cost_wepiao_distributor, profit_wepiao_distributor, distirbutor_roi, '', production_income_wepiao, copyrightIncome_wepiao, total_income_wepiao_investor, total_cost_wepiao_investor, profit_wepiao_investor, investor_roi]
# end of calculatorDataRow


# prepare data for data frame
def calculatorDataFrame(boxoffice_even_combined, boxoffice_even_distributor, boxoffice_even_investor, parameter_list, boxoffice_net_percetage, result):
    """
    Prepare for calculator data frame.
    """
    data = pd.DataFrame(columns=['boxoffice', 'boxoffice_net', 'depositCost', 'issue_income', 'proxy_income', 'net_issue_income', 'propaganda', 'production', 'profit_film', 'casts_profit', 'net_profit_film', 'copyrightIncome', 'total_income_investor', 'total_profit_investor', 'roi_film_investor', 'blank_1', 'proxy_income_wepiao_combined', 'propaganda_income_wepiao_combined', 'production_income_wepiao_combined', 'copyrightIncome_wepiao_combined', 'total_income_wepiao_combined', 'propaganda_cost_wepiao_combined',
                                 'production_cost_wepiao_combined', 'total_cost_wepiao_combined', 'profit_wepiao_combined', 'roi_wepiao_combined', 'blank_2', 'proxy_income_wepiao_distributor', 'propaganda_income_wepiao_distributor', 'total_income_wepiao_distributor', 'total_cost_wepiao_distributor', 'profit_wepiao_distributor', 'roi_wepiao_distributor', 'blank_3', 'production_income_wepiao_investor', 'copyrightIncome_wepiao_investor', 'total_income_wepiao_investor', 'total_cost_wepiao_investor', 'profit_wepiao_investor', 'roi_wepiao_investor'])

    # determine max box office and step size
    boxoffice_max = max(2 * boxoffice_even_combined, 2 * boxoffice_even_distributor,
                        2 * boxoffice_even_investor, 10 * int(parameter_list['production']))
    million_1 = 100
    if boxoffice_max < million_1 * 10:
        boxoffice_step = million_1
    elif boxoffice_max < million_1 * 100:
        boxoffice_step = million_1 * 10
    elif boxoffice_max < million_1 * 1000:
        boxoffice_step = million_1 * 20
    elif boxoffice_max < million_1 * 5000:
        boxoffice_step = million_1 * 50
    else:
        boxoffice_step = round(boxoffice_max / 1000000) * 10000
    #print(boxoffice_max, boxoffice_step)

    # start the loop
    boxoffice_start = 0

    # for wepiao as investor, distributor and combined
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_combined, parameter_list, boxoffice_net_percetage, result)
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_distributor, parameter_list, boxoffice_net_percetage, result)
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_investor, parameter_list, boxoffice_net_percetage, result)

    while boxoffice_start < boxoffice_max:
        data.loc[len(data.index)] = calculatorDataRow(
            boxoffice_start, parameter_list, boxoffice_net_percetage, result)
        # print(list(data.columns))
        #print(calculatorDataRow(boxoffice_start, parameter_list, boxoffice_net_percetage, result))
        boxoffice_start += boxoffice_step

    # now, we would introduce boxoffice when 25 percentage profit is made
    boxoffice_even_combined_25, boxoffice_even_distributor_25, boxoffice_even_investor_25 = BreakEvenBoxoffice(
        parameter_list, result, even_percentage=25)
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_combined_25, parameter_list, boxoffice_net_percetage, result)
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_distributor_25, parameter_list, boxoffice_net_percetage, result)
    data.loc[len(data.index)] = calculatorDataRow(
        boxoffice_even_investor_25, parameter_list, boxoffice_net_percetage, result)

    # sort by boxoffice
    data.sort(['boxoffice'], ascending=[True], inplace=True)
    data = data.drop_duplicates()
    data.reset_index(drop=True, inplace=True)

    return data
# end of calculatorDataFrame


# carete chart for excel files
def CreateChart(workbook, worksheet_return, sheetname, table_data_len):
    """
    Adding chart to sheet of interest.
    """
    # add chart to reflect return line
    chart = workbook.add_chart({'type': 'line'})
    #chart.add_series({'values': '=收益测算表!$A$2:$A$' + str(len(table_data.index) + 1)})
    # investors_film
    chart.add_series({
        'categories': '=' + str(sheetname) + '!$A$2:$A$' + str(table_data_len + 1),
        'values': '=' + str(sheetname) + '!$O$2:$O$' + str(table_data_len + 1),
        'line': {
            'color': 'red',
            'width': 1
        },
        'name': '影片投资方',
        'marker': {'type': 'automatic'},
    })
    # combined
    chart.add_series({
        'categories': '=' + str(sheetname) + '!$A$2:$A$' + str(table_data_len + 1),
        'values': '=' + str(sheetname) + '!$Z$2:$Z$' + str(table_data_len + 1),
        'line': {
            'color': 'red',
            'width': 1
        },
        'name': '微影(发行+投资)',
        'marker': {'type': 'automatic'},
    })
    # investor
    chart.add_series({
        'categories': '=' + str(sheetname) + '!$A$2:$A$' + str(table_data_len + 1),
        'values': '=' + str(sheetname) + '!$AN$2:$AN$' + str(table_data_len + 1),
        'line': {
            'color': 'blue',
            'width': 1
        },
        'name': '微影(仅投资)',
        'marker': {'type': 'automatic'},
    })
    # distributor
    chart.add_series({
        'categories': '=' + str(sheetname) + '!$A$2:$A$' + str(table_data_len + 1),
        'values': '=' + str(sheetname) + '!$AG$2:$AG$' + str(table_data_len + 1),
        'line': {
            'color': 'blue',
            'width': 1
        },
        'name': '微影(仅发行)',
        'marker': {'type': 'automatic'},
    })
    # set chart title
    chart.set_title({
        'name': '影片投资方和微影投资回报率vs票房',
        'name_font': {
            'color': 'blue',
        },
    })
    # axis
    chart.set_x_axis({
        'name': '影片票房（万）',
        'name_font': {
            'name': 'Courier New',
            'color': '#92D050'
        },
        'num_font': {
            'name': 'Arial',
            'color': '#00B0F0',
        },
    })
    chart.set_y_axis({
        'name': '回报率%',
        'name_font': {
            'name': 'Century',
            'color': '#92D050'
        },
        'num_font': {
            'bold': True,
            'italic': True,
            'underline': False,
            'color': '#7030A0',
        },
        'num_format': '0.0"%"',
    })
    # legend
    chart.set_legend({'font': {'bold': 1, 'italic': 1}})
    # chart size
    chart.set_size({'width': 900, 'height': 600})
    # chartarea background color
    # chart.set_chartarea({
    #    'border': {'none': True},
    #    'fill':   {'color': 'none'}
    #})
    # insert chart
    worksheet_return.insert_chart('B' + str(table_data_len + 5), chart)
# end of CreateChart


# create and write into excel files
def CreateExcel(table_data, boxoffice_even_combined, boxoffice_even_distributor, boxoffice_even_investor, json):
    """
    Create excel file and write what we want into it.
    """
    # write to excel
    filename = 'app/static/calculator/收益测算_' + datetime.date.today().strftime('%Y%m%d') + \
        '_' + str(int(round(time.time() * 1000))) + '.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # get parameter data and write it into excel
    """
    parameter_data = pd.DataFrame({
            'header': ['专项基金比例%', '增指税比例%', '分账比例%', '发行代理费比例%', '主创分红比例%', '投资方其他收入（万）', '微影发行代理费比例%', '微影其他收入（万）', '微影宣发占比%', '微影制作投资占比%', '影片宣发费用（万）', '影片制作预算（万）', '微影宣发投入（万）', '微影制作投入（万）', '影片投资方盈亏票房（万）', '微影盈亏票房（万）'], 
            'parameter': [json['zzbPercentage'], json['taxPercentage'], json['boxPercentage'], json['proxyPercentage'], json['castsPercentage'], json['copyrightIncome'], json['proxyPercentage_wepiao'], json['copyrightIncome_wepiao'], json['propagandaPercentage_wepiao'], json['productionPercentage_wepiao'], json['propaganda'], json['production'], json['propaganda_wepiao'], json['production_wepiao'], json['boxoffice_even_investor'], json['boxoffice_even_combined']]
    })
    parameter_data.to_excel(writer, sheet_name = '参数表', index = False, header = False)
    """

    # detailed data
    table_data.columns = ['总票房(万)', '净票房(万)', '影片其他成本（如押金，万）', '发行收入(万)', '发行代理费(万)', '净发行收入(万)', '宣发成本(万)', '制作成本(万)', '影片利润(万)', '主创分红(万)', '待投资方分配利润(万)', '影片其他收入(万)', '影片投资方总收入(万)', '影片投资方利润(万)', '影片投资方回报率(%)', 'blank_1', '微影发行代理收入(万)', '微影宣发回收(万)', '微影制作回收&利润(万)', '微影其他收入(万)', '微影总收入(万)',
                          '微影宣发成本(万)', '微影制作成本(万)', '微影总成本(万)', '微影利润(万)', '微影回报率(%)', 'blank_2', '微影发行代理收入(仅发行,万)', '微影宣发回收(仅发行,万)', '微影(仅发行)总收入(万)', '微影(仅发行)总成本(万)', '微影(仅发行)利润(万)', '微影(仅发行)回报率(%)', 'blank_3', '微影制作回收&利润(仅投资,万)', '微影其他收入(仅投资,万)', '微影(仅投资)总收入(万)', '微影(仅投资)总成本(万)', '微影(仅投资)利润(万)', '微影(仅投资)回报率(%)']
    # Updated at 14:17 on Oct 18th, 2016:
    # Actually, we invoke a bug here because we name three blank column, and thus when we write into excel file, once it encounters a blank column name, it would expand into three columns. It is really a small bug, but finally we find it.
    # write into excel
    table_data.to_excel(writer, sheet_name='收益测算表', index=False, header=True)

    # after saving, now we can do some formatting
    workbook = writer.book
    # general formatting
    number_format = workbook.add_format(
        {'align': 'right', 'num_format': '#,##0', 'bold': True, 'bottom': 0})
    # Total percent format
    percent_format = workbook.add_format(
        {'align': 'right', 'num_format': '0.00', 'bold': True, 'bottom': 0})
    # text_format
    text_format = workbook.add_format(
        {'align': 'left', 'bold': True, 'bottom': 0})
    # add color
    negative_color = workbook.add_format(
        {'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    positive_color = workbook.add_format(
        {'bg_color': '#C6EFCE', 'font_color': '#006100'})
    # set header
    header_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10,
                                      'bold': True, 'bg_color': '#20DF2B', 'align': 'center', 'color': 'white'})

    # 收益测算表
    worksheet_return = writer.sheets['收益测算表']
    worksheet_return.set_zoom(90)

    # font size
    workbook_format = workbook.add_format()
    workbook_format.set_font_size(10)

    # set column
    worksheet_return.set_column('A:N', 20, number_format)
    worksheet_return.set_column('O:O', 20, percent_format)
    worksheet_return.set_column('Q:Y', 20, number_format)
    worksheet_return.set_column('Z:Z', 20, percent_format)
    worksheet_return.set_column('AB:AF', 20, number_format)
    worksheet_return.set_column('AG:AG', 20, percent_format)
    worksheet_return.set_column('AI:AM', 20, number_format)
    worksheet_return.set_column('AN:AN', 20, percent_format)

    # for blank columns' header
    worksheet_return.write('P1', '')
    worksheet_return.write('AA1', '')
    worksheet_return.write('AH1', '')

    # set header format
    #header_fmt = workbook.add_format({'font_name': 'Arial', 'font_size': 10, 'bold': True})
    worksheet_return.set_row(0, None, header_fmt)

    # for color
    worksheet_return.conditional_format('N2:O' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '>', 'value': 0, 'format': positive_color})
    worksheet_return.conditional_format('N2:O' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '<', 'value': 0, 'format': negative_color})
    worksheet_return.conditional_format('Y2:Z' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '>', 'value': 0, 'format': positive_color})
    worksheet_return.conditional_format('Y2:Z' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '<', 'value': 0, 'format': negative_color})
    worksheet_return.conditional_format('AF2:AG' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '>', 'value': 0, 'format': positive_color})
    worksheet_return.conditional_format('AF2:AG' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '<', 'value': 0, 'format': negative_color})
    worksheet_return.conditional_format('AM2:AN' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '>', 'value': 0, 'format': positive_color})
    worksheet_return.conditional_format('AM2:AN' + str(len(table_data.index) + 1), {
                                        'type': 'cell', 'criteria': '<', 'value': 0, 'format': negative_color})

    #worksheet_return.set_row(0, None, header_fmt)
    #worksheet_return.set_row(0, len(table_data.index), 20)
    #worksheet_return.set_row(0, 10, format = {'font_size': 10})
    # for break-even comment and color
    # find the position of break-even boxoffice
    position_investor = table_data[(table_data['总票房(万)'] > boxoffice_even_investor - 1) & (
        table_data['总票房(万)'] < boxoffice_even_investor + 1)].index[0] + 2
    position_combined = table_data[(table_data['总票房(万)'] > boxoffice_even_combined - 1) & (
        table_data['总票房(万)'] < boxoffice_even_combined + 1)].index[0] + 2
    position_distributor = table_data[(table_data['总票房(万)'] > boxoffice_even_distributor - 1) & (
        table_data['总票房(万)'] < boxoffice_even_distributor + 1)].index[0] + 2
    #print(table_data[(table_data['总票房(万)'] > boxoffice_even_film - 1) & (table_data['总票房(万)'] < boxoffice_even_film + 1)], table_data[(table_data['总票房(万)'] > boxoffice_even_wepiao - 1) & (table_data['总票房(万)'] < boxoffice_even_wepiao + 1)])
    #print(position_film, position_wepiao)
    worksheet_return.write_comment('A' + str(position_investor), '影片（仅投资）盈亏票房')
    worksheet_return.write_comment(
        'A' + str(position_combined), '微影（投资+发行）盈亏票房')
    worksheet_return.write_comment(
        'A' + str(position_distributor), '微影（仅发行）盈亏票房')

    # insert line into sheet of interest
    CreateChart(workbook, worksheet_return, '收益测算表', len(table_data.index))

    # 参数表
    """
    worksheet_parameter = writer.sheets['参数表']
    worksheet_parameter.set_zoom(90)
    #set column
    worksheet_parameter.set_column('A:A', 30, text_format)
    worksheet_parameter.set_column('B:B', 20, number_format)
    
    #add another sheet to work with excel formulas
    worksheet_formula = workbook.add_worksheet('收益测算_有公式')
    
    #we have to write all boxoffice values into this sheet
    for row_i in range(len(table_data.index)):
        row_number = row_i + 2
        #worksheet_formula.set_header('&CHere is some centered text.')
        #worksheet_formula.write('A' + str(row_number), boxoffice_even_wepiao)
        worksheet_formula.write('A' + str(row_number), table_data['总票房(万)'][row_i])
        #worksheet_formula.write_formula('B2', '=A2*2') 

        #1: boxoffice_net
        worksheet_formula.write_formula('B' + str(row_number), '=(1-参数表!$B$1/100-参数表!$B$2/100)*A' + str(row_number))
        #2: issue_income
        worksheet_formula.write_formula('C' + str(row_number), '=(参数表!$B$3/100)*B' + str(row_number))
        #3: proxy
        worksheet_formula.write_formula('D' + str(row_number), '=(参数表!$B$4/100)*C' + str(row_number))
        #4: issue_net
        worksheet_formula.write_formula('E' + str(row_number), '=C' + str(row_number) + '-D' + str(row_number))
        #5: propaganda_film
        worksheet_formula.write_formula('F' + str(row_number), '=if(E' + str(row_number) +'<参数表!$B$11, E' + str(row_number) + ', 参数表!$B$11)')
        #6: production_film
        worksheet_formula.write_formula('G' + str(row_number), '=min(参数表!$B$12, max(0, E' + str(row_number) + '-F' + str(row_number) + '))')
        #7: casts_income
        worksheet_formula.write_formula('H' + str(row_number), '=max(0, 参数表!$B$5*(E' + str(row_number) + '-F' + str(row_number) + '-G' + str(row_number) + '))/100')
        #8: copyrightIncome
        worksheet_formula.write_formula('I' + str(row_number), '=参数表!$B$6')
        #9: investor_income
        worksheet_formula.write_formula('J' + str(row_number), '=E' + str(row_number) + '-F' + str(row_number) + '-H' + str(row_number) + '+I' + str(row_number))
        #10: investor_profit
        worksheet_formula.write_formula('K' + str(row_number), '=J' + str(row_number) + '-参数表!$B$12')
        #11: investor_roi
        worksheet_formula.write_formula('L' + str(row_number), '=K' + str(row_number) + '*100/参数表!$B$12')
        #12: blank
        worksheet_formula.write_formula('M' + str(row_number), '=""')
        #13: proxy_wepiao
        worksheet_formula.write_formula('N' + str(row_number), '=(参数表!$B$7/100)*C' + str(row_number))
        #14: propaganda_wepiao
        worksheet_formula.write_formula('O' + str(row_number), '=(参数表!$B$9/100)*F' + str(row_number))
        #15: production_wepiao
        worksheet_formula.write_formula('P' + str(row_number), '=max(0, E' + str(row_number) + '-F' + str(row_number) + '-H' + str(row_number) + ')*参数表!$B$10/100')
        #16: copyrightIncome_wepiao
        worksheet_formula.write_formula('Q' + str(row_number), '=参数表!$B$8')
        #17: income_total_wepiao
        worksheet_formula.write_formula('R' + str(row_number), '=N' + str(row_number) + '+O' + str(row_number) + '+P' + str(row_number) + '+Q' + str(row_number))
        #18: cost_total_wepiao
        worksheet_formula.write_formula('S' + str(row_number), '=参数表!$B$13+参数表!$B$14')
        #19: profit_wepiao
        worksheet_formula.write_formula('T' + str(row_number), '=R' + str(row_number) + '-S' + str(row_number))
        #20: wepiao_roi
        worksheet_formula.write_formula('U' + str(row_number), '=T' + str(row_number) + '*100/S' + str(row_number))
    
    #set formatting for 收益测算(有公式)
    worksheet_formula.set_column('A:K', 20, number_format)
    worksheet_formula.set_column('L:L', 20, percent_format)
    worksheet_formula.set_column('N:T', 20, number_format)
    worksheet_formula.set_column('U:U', 20, percent_format)
    
    #copy header from sheet 2
    for column in list(string.ascii_uppercase):
        worksheet_formula.write_formula(str(column) + '1', '=收益测算表!$'+str(column)+'$1&""')
    #end of this worksheet section
    
    #just for testing
    #worksheet_formula.write_formula('A100', '=收益测算_有公式!$A$1')
    
    #add chart to the third sheet
    #insert line into sheet of interest
    CreateChart(workbook, worksheet_formula, '收益测算_有公式', len(table_data.index))
    """

    # save and close the excel writer
    writer.save()
    writer.close()

    return filename
# end of CreateExcel


# convert json objects into float
def Json2Float(json):
    """
    Convert each element of a json into float except specifically pointed out.
    """
    # convert json objects to float
    for element_i in list(json.keys()):
        if element_i in ['production_wepiao_checkbox', 'propaganda_wepiao_checkbox', 'copyrightIncome_wepiao_checkbox']:
            continue
        json[element_i] = float(json[element_i])

    return json
# end of Json2Float


# calculator processing
@calculator.route('/calculatorProcessing', methods=['POST'])
def calculatorProcessing():
    """
    Process the calculator data, and return with json.
    """
    result = request.json
    # convert every element of result into float
    for element_i in result.keys():
        result[element_i] = Json2Float(result[element_i])

    # main form
    json = result['main']

    # firstly, we calculate break-even boxoffice for film and wepiao,
    # respectively
    boxoffice_even_distributor, boxoffice_even_investor, boxoffice_even_combined = BreakEvenBoxoffice(
        json, result)
    json['boxoffice_even_distributor'] = boxoffice_even_distributor
    json['boxoffice_even_investor'] = boxoffice_even_investor
    json['boxoffice_even_combined'] = boxoffice_even_combined

    # get pd data frame
    boxoffice_net_percetage = (
        1 - json['taxPercentage'] / 100 - json['zzbPercentage'] / 100)
    table_data = calculatorDataFrame(boxoffice_even_combined, boxoffice_even_distributor,
                                     boxoffice_even_investor, json, boxoffice_net_percetage, result)
    # write table_data into json
    json['data'] = table_data.to_dict('records')
    # print(list(table_data.columns))

    # write into excel file and return filename
    filename = CreateExcel(table_data, boxoffice_even_combined,
                           boxoffice_even_distributor, boxoffice_even_investor, json)

    # save file path into json
    json['filename'] = filename[3:]
    # print(jsonify(json))

    return jsonify(json)
# end of calculatorProcessing
