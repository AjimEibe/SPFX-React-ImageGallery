import * as React from 'react';
import styles from './Carousel.module.scss';
import { ICarouselProps } from './ICarouselProps';
import { ICarouselState } from './ICarouselState';
import { RandomIndex } from './RandomIndex';
import { escape } from '@microsoft/sp-lodash-subset';
import spservices from '../../../spservices/spservices';
import * as microsoftTeams from '@microsoft/teams-js';
import { ICarouselImages } from './ICarouselmages';
import 'video-react/dist/video-react.css'; // import css
import { Player, BigPlayButton } from 'video-react';
import Slider from "react-slick";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import * as $ from 'jquery';
import { FontSizes, } from '@uifabric/fluent-theme/lib/fluent/FluentType';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import * as strings from 'CarouselWebPartStrings';
import { DisplayMode } from '@microsoft/sp-core-library';
import { CommunicationColors } from '@uifabric/fluent-theme/lib/fluent/FluentColors';
import {
	Spinner,
	SpinnerSize,
	MessageBar,
	MessageBarType,
	Label,
	Icon,
	ImageFit,
	Image,
	ImageLoadState,
} from 'office-ui-fabric-react';


export default class Carousel extends React.Component<ICarouselProps, ICarouselState> {
	private spService: spservices = null;
	private _teamsContext: microsoftTeams.Context = null;
	public rnd: RandomIndex = null;

	public constructor(props: ICarouselProps) {
		super(props);
		this.spService = new spservices(this.props.context);

		this.rnd = new RandomIndex(0);


		this.state = {
			isLoading: false,
			errorMessage: '',
			hasError: false,
			teamsTheme: 'default',
			photoIndex: 0,
			carouselImages: [],
			loadingImage: true
		};
	}


	private onConfigure() {
		// Context of the web part
		this.props.context.propertyPane.open();
	}


	private async loadPictures() {

		const tenantUrl = `https://${location.host}`;
		let galleryImages: ICarouselImages[] = [];
		let carouselImages: React.ReactElement<HTMLElement>[] = [];

		try {
			const images = await this.spService.getImages(this.props.siteUrl, this.props.list, this.props.numberImages);
			//	this.numberImages = this.props.numberImages;

			for (const image of images) {

				if (image.FileSystemObjectType == 1) continue; // by pass folder item
				const pURL = `${tenantUrl}/_api/v2.0/sharePoint:${image.File.ServerRelativeUrl}:/driveItem/thumbnails/0/large/content?preferNoRedirect=true `;
				const thumbnailUrl = `${tenantUrl}/_api/v2.0/sharePoint:${image.File.ServerRelativeUrl}:/driveItem/thumbnails/0/c240x240/content?preferNoRedirect=true `;

				let mediaType: string = '';
				switch (image.File_x0020_Type) {
					case 'jpg':
					case 'jpeg':
					case 'png':
					case 'tiff':
					case 'jfif':
					case 'gif':
						mediaType = 'image';
						break;
					case 'mp4':
						mediaType = 'video';
						break;
					default:
						continue;
						break;
				}

				galleryImages.push(
					{
						imageUrl: pURL,
						mediaType: mediaType,
						serverRelativeUrl: image.File.ServerRelativeUrl,
						caption: image.Title ? image.Title : image.File.Name,
						description: image.Description ? image.Description : '',
						linkUrl: ''
					},
				);

				// Create Gallery Slides from Images


				carouselImages = galleryImages.map((galleryImage, i) => {
					return (
						<div className='slideLoading'>
							<div>
								<Image src={galleryImage.imageUrl}
									onLoadingStateChange={async (loadState: ImageLoadState) => {
										console.log('imageload Status ' + i, loadState, galleryImage.imageUrl);
										if (loadState == ImageLoadState.loaded) {
											this.setState({ loadingImage: false });
										}
									}}
									height={'400px'}

								/>
							</div>
						</div>
					);
				}
				);
				this.setState({ carouselImages: carouselImages, isLoading: false });

			}
		} catch (error) {
			this.setState({ hasError: true, errorMessage: decodeURIComponent(error.message) });
		}
	}



	public async componentDidMount() {
		await this.loadPictures();

	}

	public async componentDidUpdate(prevProps: ICarouselProps) {

		if (!this.props.list || !this.props.siteUrl) return;
		// Get  Properties change
		if (prevProps.list !== this.props.list || prevProps.numberImages !== this.props.numberImages) {
			await this.loadPictures();
		}
	}


	public changeImage = () => {
		let rndval: number;
		var min = 0;
		var max = this.state.carouselImages.length;
		rndval = min + (Math.random() * (max - min));
		var y: number = +rndval.toFixed();
		this.setState({ photoIndex: y });
		this.setState({ photoIndex: y });
		this.setState({ photoIndex: y });
		//this.rnd.random = 2;
	}


	public render(): React.ReactElement<ICarouselImages> {
		console.log("Random value");
		console.log(this.state.photoIndex);
		return (
			<div>
				<div>
					{
						<div style= {{display: 'flex', alignItems:'center', justifyContent:'center'}}>
							{
								this.state.carouselImages[this.state.photoIndex]
							}
						</div>
					}
				</div>
				<div style= {{display: 'flex', alignItems:'center', justifyContent:'center'}}>
				<br></br><br></br>
					<button onClick={this.changeImage}> Next Image</button>
				</div>
			</div>
		);
	}
}