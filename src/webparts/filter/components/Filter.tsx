import * as React from 'react';
import styles from './Filter.module.scss';
import { IFilterProps } from './IFilterProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';


export default class Filter extends React.Component<IFilterProps, {}> {
	list: any;
	private async onGetListItemsClicked(event: React.MouseEvent<HTMLButtonElement>): Promise<void> {
		event.preventDefault();

		this.list = await this.getListItems();
		console.log(this.list);

	}

	public async getListItems() {
		const response = await this.props.context.spHttpClient.get(
			`https://elfodev.sharepoint.com/_api/web/lists`,
			SPHttpClient.configurations.v1);
		return (await response.json());
	}

	public render(): React.ReactElement<IFilterProps> {

		const {
			context,
			description,
			isDarkTheme,
			environmentMessage,
			hasTeamsContext,
			userDisplayName
		} = this.props;

		return (
			<section className={`${styles.filter} ${hasTeamsContext ? styles.teams : ''}`}>
				<div className={styles.welcome}>
					<img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
					<h2>Well done, {escape(userDisplayName)}!</h2>
					<div>{environmentMessage}</div>
					<div>Web part property value: <strong>{escape(description)}</strong></div>
				</div>
				<div>
					<h3>Welcome to SharePoint Framework!</h3>
					<p>
						The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
					</p>
					<h4>Learn more about SPFx development:</h4>

				</div>
				<button type="button" onClick={this.onGetListItemsClicked}>LoadLists</button>

			</section>
		);

	}
} 
