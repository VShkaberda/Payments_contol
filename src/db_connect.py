# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 09:43:48 2018

@author: v.shkaberda
"""
from functools import wraps
from tkPayments import NetworkError
import pyodbc

def monitor_network_state(method):
    """ Show error message in case of network error.
    """
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except pyodbc.Error as e:
            # Network error
            if e.args[0] in ('01000', '08S01', '08001'):
                NetworkError()
    return wrapper


class DBConnect(object):
    """ Provides connection to database and functions to work with server.
    """
    def __init__(self, *, server, db):
        self._server = server
        self._db = db

    def __enter__(self):
        # Connection properties
        conn_str = (
            'Driver={{SQL Server}};'
            'Server={0};'
            'Database={1};'
            'Trusted_Connection=yes;'.format(self._server, self._db)
            )
        self.__db = pyodbc.connect(conn_str)
        self.__cursor = self.__db.cursor()
        return self

    def __exit__(self, type, value, traceback):
        self.__db.close()

    @monitor_network_state
    def access_check(self):
        """ Check user permission.
            If access prmitted returns True, otherwise None.
        """
        self.__cursor.execute("exec [payment].[Access_Check]")
        access = self.__cursor.fetchone()
        # check AccessType and isSuperUser
        if access and (access[0] in (1, 2) or access[1]):
            return True

    @monitor_network_state
    def create_request(self, userID, mvz, office, categoryID, contragent, csp,
                       plan_date, sumtotal, nds, text, approval):
        """ Executes procedure that creates new request.
        """
        query = '''
        exec payment.create_request @UserID = ?,
                                    @MVZ = ?,
                                    @Office = ?,
                                    @CategoryID = ?,
                                    @Contragent = ?,
                                    @date_planed = ?,
                                    @Description = ?,
                                    @SumNoTax = ?,
                                    @Tax = ?,
                                    @CSP = ?,
                                    @Approval = ?
            '''
        try:
            self.__cursor.execute(query, userID, mvz, office, categoryID,
                                  contragent, plan_date, text,
                                  sumtotal, nds, csp, approval)
            request_allowed = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_allowed
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def get_user_info(self):
        """ Returns information about current user based on ORIGINAL_LOGIN().
        """
        self.__cursor.execute("select UserID, ShortUserName, AccessType, isSuperUser \
          from payment.People \
          where UserLogin = right(ORIGINAL_LOGIN(), len(ORIGINAL_LOGIN()) - charindex( '\\' , ORIGINAL_LOGIN()))")
        return self.__cursor.fetchone()

    @monitor_network_state
    def get_allowed_initiators(self, UserID, AccessType, isSuperUser):
        """ Determines list of persons that should be shown in filters.
        """
        query = '''
        exec payment.get_allowed_initiators @UserID = ?,
                                            @AccessType = ?,
                                            @isSuperUser = ?
        '''
        self.__cursor.execute(query, UserID, AccessType, isSuperUser)
        return [(None, 'Все'),] + self.__cursor.fetchall()

    @monitor_network_state
    def get_approvals(self, paymentID):
        """ Returns all approvals of the request with id = paymentID.
        """
        query = "exec payment.get_approvals @paymentID = ?"
        self.__cursor.execute(query, paymentID)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_approvals_for_first_stage(self):
        """ Returns all approvals for first stage who can be chosen.
        """
        query = "exec payment.get_approvals_for_first_stage"
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_categories(self, user_info):
        """ Returns list of available MVZ for current user.
        """
        query = "exec payment.get_categories @isSuperUser = ?"
        self.__cursor.execute(query, user_info.isSuperUser)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_limit_for_month_by_date(self, UserID, date):
        """ Returns remaining limit for UserID for month corresponding to
            month of specified date.
        """
        query = '''
        exec payment.get_limit_for_month_by_date @UserID = ?, @date = ?
        '''
        self.__cursor.execute(query, UserID, date)
        return self.__cursor.fetchone()[0]

    @monitor_network_state
    def get_limits_info(self):
        """ Returns info about limits for all users.
        """
        query = '''
        select UserID, UserName,
               cast(userCreateRequestLimit as float) as userCreateRequestLimit,
               resetCreateRequestLimit
        from payment.People
        where AccessType in (1, 2)
            or isSuperUser = 1
        order by UserName
            '''
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_MVZ(self, user_info):
        """ Returns list of available MVZ for current user.
        """
        query = '''
        exec payment.get_MVZ @UserID = ?,
                             @AccessType = ?,
                             @isSuperUser = ?
        '''
        self.__cursor.execute(query, user_info.UserID, user_info.AccessType,
                              user_info.isSuperUser)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_paymentslist(self, user_info, initiator, mvz, office,
                         plan_date_m, plan_date_y, sumtotal_from, sumtotal_to,
                         nds, just_for_approval):
        """ Generates query according to user's acces type and filters.
        """
        query = '''
        select pl.ID as ID, pl.UserID as InitiatorID,
           'ЛГ-' + replace(convert(varchar, date_created, 102),'.','') + '_' + cast(pl.ID as varchar(7)) as Num,
           pp.ShortUserName, cast(date_created as date) as date_created,
           cast(date_created as smalldatetime) as datetime_created,
           isnull(CSP, '') as CSP, isnull(obj.MVZsap, '') as MVZsap,
           isnull(co.FullName, 'ТехМВЗ') as FullName, obj.ServiceName,
           cat.CategoryName, isnull(Contragent, '') as Contragent, date_planed,
           SumNoTax, cast(SumNoTax * ((100 + Tax) / 100.0) as numeric(11, 2)),
           p.ValueName as StatusName, p.ValueDescription, pl.Description, appr.UserID,
           case when pl.StatusID = 1 then isnull(pappr.ShortUserName, '') else '' end as approval
        from payment.PaymentsList pl
        join payment.ListObjects obj on pl.ObjectID = obj.ID
        join payment.ListCategories cat on pl.CategoryID = cat.ID
        left join LogisticFinance.BTool.aid_CostObject_Detail co on co.SAPMVZ = obj.MVZsap
        join payment.People pp on pl.UserID = pp.UserID
        join LogisticFinance.dbo.GlobalParamsLines p on pl.StatusID = p.idParamsLines
                                    and p.idParams = 2
                                    and p.Enabled = 1
        -- approvals
        left join payment.PaymentsApproval appr on pl.ID = appr.PaymentID
                                           and appr.is_active_approval = 1
        left join payment.People pappr on appr.UserID = pappr.UserID
        where 1=1
        '''
        if just_for_approval:
            if user_info.UserID == 24:
                query += 'and pl.StatusID = 1 and appr.UserID in (9, 24)\n'
            else:
                query += 'and pl.StatusID = 1 and appr.UserID = {}\n'.format(user_info.UserID)
        else:
            if not user_info.isSuperUser:
                query += "and (pl.UserID = {0} or exists(select * from payment.PaymentsApproval _appr \
                        where pl.ID = _appr.PaymentID and _appr.UserID = {0}))\n".format(user_info.UserID)
            if initiator:
                query += "and pl.UserID = {}\n".format(initiator)
            if mvz:
                query += "and obj.MVZsap = '{}'\n".format(mvz)
            if office:
                query += "and obj.ServiceName = '{}'\n".format(office)
            if plan_date_y and all(map(str.isdigit, str(plan_date_y))):
                query += "and year(date_planed) = {}\n".format(plan_date_y)
            if plan_date_m:
                query += "and month(date_planed) in ({})\n".format(plan_date_m)
            if sumtotal_from:
                query += "and SumNoTax >= {}\n".format(sumtotal_from)
            if sumtotal_to:
                query += "and SumNoTax <= {}\n".format(sumtotal_to)
            if not nds == -1:
                query += "and Tax = {}\n".format(nds)
        # specific sorting for several people
        if user_info.UserID in (42, 81, 75):
            query += "order by IIF(pl.StatusID in (2, 4), 2, 1) ASC, ID DESC"
        else:
            query += "order by ID DESC" # the same as created(datetime) DESC
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def raw_query(self, query):
        """ Takes the query and returns output from db.
        """
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def update_confirmed(self, userID, paymentID, is_approved):
        """ Adds information to db if paymentID has been approved.
        """
        query = '''
        exec payment.approve_request @UserID = ?,
                                     @paymentID = ?,
                                     @is_approved = ?
        '''
        self.__cursor.execute(query, userID, paymentID, is_approved)
        self.__db.commit()

    @monitor_network_state
    def update_discarded(self, discardID):
        """ Set status of request to "discarded".
        """
        query = "exec payment.discard_request @paymentID = ?"
        self.__cursor.execute(query, discardID)
        self.__db.commit()

    @monitor_network_state
    def update_limits(self, limits):
        """ Update user limits.
        """
        query = '''
        UPDATE payment.People
        SET resetCreateRequestLimit = ?, userCreateRequestLimit = ?
        where UserID = ?
        '''
        limits = [tuple(reversed(info)) for info in limits]
        try:
            self.__cursor.executemany(query, limits)
            self.__db.commit()
            return 1
        except pyodbc.ProgrammingError:
            return 0


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s64', db='CB') as sql:
        query = 'select 42'
        assert sql.raw_query(query)[0][0] == 42, 'Server returns no output.'
    print('Connected successfully.')
    input('Press Enter to exit...')
