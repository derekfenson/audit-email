
from abc import ABC, abstractmethod
from utilities import send_mail_via_com

class AuditEmailer(ABC):
	"""Abstract class for creating and sending emails based on audit report of a project workflow
	Takes as an input a dataframe unique by study/investigation and deliverable date, with variables indicating specific erros with project workflow
	Create specific HTML email body for each study/investigation, and send the email from 
	"""

	""" Properties """
	@property
	@abstractmethod
	def audit_report_path(self):
		"""Path to the audit report"""
		raise NotImplemented

	@property
	@abstractmethod
	def output_path(self):
		"""Path where the email .msg files will be saved, as well as the final csv report"""
		raise NotImplemented

	@property
	@abstractmethod
	def study_prefix(self):
		"""String containing the prefix used for naming studies for this project (ex. study, inv, etc.)"""
		raise NotImplemented

	@abstractmethod
	@property
	def sender(self):
		"""String containing email address for the sender of the emails"""
		raise NotImplemented

	@property
	def inv_prefix(self):
		"""String containing the prefix used for naming investigations for this project (ex. inv, request, etc.)
		Allows for a multi-level organiziation for investigations, but by default assumes there is only one level and returns None by default
		"""
		return None

	def __init__(self):
		super().__init__()

	def __call__(self, run_date, autosend):
		"""Main caller for creating and sending emails
		"""
		study_df=read_csv(r"{0}\Audit_Report_{1}.csv".format(self.audit_report_path,run_date), dtype=str)
		study_df=study_df.applymap(lambda x: "" if x.lower()=="nan" else x)
		study_df=study_df.apply(patch, axis=1)
		self.send(study_df, run_date, autosend)


	def send(self, study_df, run_date, autosend=False):
		"""Create, save, and send emails for this project
		study_df is a dataframe containing all meta data for the project. Must contain the following variables:
			deliver_error, data_error, log_error, svn_error, study, (inv is optional), date, deliver_type
		run_date is the date of the run, in yyyy_mm_dd format
		autosend indicates if the emails should be automatically sent
		"""
		group_vars=self.study_prefix
		if self.inv_prefix is not None:
			group_vars=[self.study_prefix,self.inv_prefix]
		study_groups=study_df.groupby(group_vars).groups
		specific_df=study_df.loc[study_df["date"] != "_ALL_",:]
		general_df=study_df.loc[study_df["date"] == "_ALL_",:]
		for k,v in study_groups:
			if type(k)==str:
				study=k
				inv=None
				subset_func=lambda x: x[self.study_prefix]==study
			else:
				study=k[0]
				inv=k[1]
				subset_func=lambda x: x[self.study_prefix]==study and x[self.inv_prefix]==inv
			specific_df=specific_df.loc[specific_df.apply(subset_func, axis=1),:]
			general_df=general_df.loc[general_df.apply(subset_func, axis=1),:]
			analyst_name=general_df.loc[0,"fullname"]
			analyst_email=general_df.loc[0,"email"]
			full_study_name=self._full_study_name(study, inv)
			name, email, body=construct_email(analyst_name, analyst_email,
												full_study_name, specific_df, general_df)
			if body != "":
				send_mail_via_com(body=html_body, 
								subject="[ACTION REQUIRED] - {0}".format(full_study_name), 
								recipient=analyst_email, 
								on_behalf_of=self.sender,
								save_path=self.define_outpath(run_date),
								autosend=autosend                      
								)

	def define_outpath(self, run_date):
		"""Create output path based on run_date if it does not exits, and return the path as a string"""
		outpath=r"{0}\{1}".format(self.output_path, run_date)
		if not os.path.exists(outpath):
			os.mkdir(outpath)
		return outpath


	def construct_email(analyst_name, analyst_email, study, inv, specific_df, general_df):
		"""Returns tuple of analyst_name, analyst_email, html_body where HTML body is the full email body for the given study/investigation

		Parameters:
		analyst_name - full name of the analyst responsibile for the study/investigation
		analyst_email - email address of the analyst responsibile for the study/investigation
		study - string containing the study number (top level number)
		inv - string containing the inv number (lower level number)
		specific_df - Dataframe containing variables indicating what errors were made spefically for each study/inv, deliverable date combination
		general_df - Dataframe contatining variables indicating what errors were made spefically for each study/inv overall. Dataframe is assumed to be one row.
		"""
		full_study_name=_full_study_name(study, inv)
		deliver_errors=[s for s in overall_df.loc[0, "deliver_error"].split("#\n") if s != ""]
		data_errors=[s for s in overall_df.loc[0, "data_error"].split("#\n") if s != ""]
		log_errors=[s for s in overall_df.loc[0, "log_error"].split("#\n") if s != ""]
		general_error=general_error_body(deliver_error_body(deliver_errors),
											data_error_body(data_errors),
											logs_error_body(log_errors))
		specific_errors=""
		for r in range(0, specific_df.shape[0]):
			deliverable="%s - %s" % (specific_df.date.iloc[r], specific_df.deliver_type.iloc[r].title())
			svn_errors=[d for d in specific_df.loc[r, "svn_error"].split("#\n") if d != ""]
			deliver_errors=[d for d in specific_df.loc[r, "deliver_error"].split("#\n") if d != ""]
			data_errors=[d for d in specific_df.loc[r, "data_error"].split("#\n") if d != ""]
			log_errors=[d for d in specific_df.loc[r, "log_error"].split("#\n") if d != ""]
			specific_errors+=specific_error_list(deliverable,
												svn_error_body(svn_errors),
												deliver_error_body(deliver_errors),
												data_error_body(data_errors),
												logs_error_body(log_errors))
		specific_error=specific_error_body(specific_errors)
		return (analyst_name.replace(" ", "_"), analyst_email, email_body(analyst_name, full_study_name, general_error, specific_error))

	def email_body(analyst_name, full_study_name, general_error_body, specific_error_body):
		"""Return the entire HTML body of the email"""
		if general_error_body != "" or specific_error_body != "":
			return """\
			<html>
				<head>
				</head>
				<body style="font: 11pt calibri;">
					Hello {0},<br>
					<p>
						We are writing to inform you that we have detected some errors in folder organization and/or file format for {1}: 
						{2}
						{3}
					</p>
					<p>
						Please read through these issues and resolve them as soon as possible. If you have any questions, please contact {4}.
					</p>

					Thanks,<br>
					<b>{4}</b>
				</body>
			</html>
		""".format(analyst_name, full_study_name, general_error_body, specific_error_body, sender_display_name)
		else:
			return ""

	def general_error_body(general_deliver_error, general_data_error, general_logs_error):
		"""Return the HTML snippet containing the errors general to the study/inv"""
		if any([item != "" for item in [general_deliver_error, general_data_error, general_logs_error]]):
		return """\
			<p>
			General:
				<ol>
					{0}
					{1}
					{2}               
				</ol>
			</p>
		""".format(general_deliver_error, general_data_error, general_logs_error)
		else:
		return ""

	def specific_error_body(deliverables_errors):
		"""Return the HTML snippet conatining the errors specific to all deliverables"""
		if deliverables_errors != "":
			return """\
			<p>
				Specific deliverables:
				<ul>
					{0}
				</ul>
			</p>
			""".format(deliverables_errors)
			else:
			return ""

		def specific_error_list(deliverable, svn_error, deliver_error, data_error, logs_error):
		"""Return the HTML snippet containing the list item conatining specific errors related to the deliverable data"""
		if any([item !="" for item in [svn_error, deliver_error, data_error, logs_error]]):
			return """\
			<li>{0}</li>  
				<ul>
					{1}
					{2}
					{3}
					{4}               
				</ul>
			""".format(deliverable, svn_error, deliver_error, data_error, logs_error)
			else:
		return ""

	def deliver_error_body(errors):
		"""Return the HTML snippet for errors in deliverables"""
	    return _prefix_error_body("Deliverables:", errors)

	def data_error_body(errors):
		"""Return the HTML snippet for errors in output data"""
	    return _prefix_error_body("Data:", errors)

	def logs_error_body(errors):
		"""Return the HTML snippet for errors in logs"""
	    return _prefix_error_body("Logs:", errors)

	def svn_error_body(errors):
		"""Return the HTML snippet for errors in the svn commit message"""
	    return _prefix_error_body("SVN:", errors)    

	def _prefix_error_body(prefix, errors):
		"""Return a string of HTML code containing a list item of an unorder list of elements of errors"""
		error_string=""
		if len(errors)>0:
			if len(errors)>1:
				error_body=["<ul>"] + ["<li>%s</li>" % e for e in errors] + ["</ul>"]
			error_string="<li>{0} {1}</li>".format(prefix, "".join(error_body))
		return error_list

	def _full_study_name(study, inv):
		"""Return a string containing full name with prefixes for the given study/inv combo"""
		full_name=" ".join([self.study_prefix.title(), study])
		if self.inv_prefix is not None:
			full_name=" ".join([full_name, self.inv_prefix.title(), inv])

	def _patch(row):
		"""Return a Series object representing a row in the study df dataframe
		Based on study, inv, and date, keep errors intact or remove errors
		row is assumed to be a row from the audit report dataframe
		"""
		return row

		
